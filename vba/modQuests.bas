Attribute VB_Name = "modQuests"
'===============================================================
' modQuests — Quest Tracking System
' Damned Moon VBA RPG Engine — Phase 4
'===============================================================
' Full quest lifecycle: start, advance, complete, fail quests.
' Quests progress through numbered stages, each with description
' text, objectives, XP rewards, and trigger effects.
'
' Data sources:
'   tbl_Quests      — quest definitions and current state
'   tbl_QuestStages — stage progression per quest
'   tbl_JournalEntries — narrative journal log
'
' tbl_Quests columns:
'   A: QuestID    B: Name       C: Type (MAIN, SIDE, HIDDEN)
'   D: Status     E: StageDesc  F: CurrentStage
'   G: TotalStages H: RewardEffects
'
' tbl_QuestStages columns:
'   A: QuestID    B: StageNum   C: StageName
'   D: Description  E: Objective  F: Requirements
'   G: OnCompleteEffects  H: XPReward
'
' Quest statuses: INACTIVE, ACTIVE, COMPLETED, FAILED
'===============================================================

Option Explicit

' ── QUEST TABLE COLUMN INDICES ──
Private Const QST_COL_ID As Long = 1           ' A: QuestID
Private Const QST_COL_NAME As Long = 2         ' B: Display name
Private Const QST_COL_TYPE As Long = 3         ' C: Type (MAIN, SIDE, HIDDEN)
Private Const QST_COL_STATUS As Long = 4       ' D: Status
Private Const QST_COL_STAGEDESC As Long = 5    ' E: Current stage description
Private Const QST_COL_CURSTAGE As Long = 6     ' F: Current stage number
Private Const QST_COL_TOTALSTAGES As Long = 7  ' G: Total stages
Private Const QST_COL_REWARD As Long = 8       ' H: Completion reward effects

' ── QUEST STAGE TABLE COLUMN INDICES ──
Private Const QSS_COL_QUESTID As Long = 1      ' A: QuestID
Private Const QSS_COL_STAGENUM As Long = 2     ' B: Stage number
Private Const QSS_COL_STAGENAME As Long = 3    ' C: Stage name
Private Const QSS_COL_DESC As Long = 4         ' D: Stage description
Private Const QSS_COL_OBJECTIVE As Long = 5    ' E: Objective text
Private Const QSS_COL_REQS As Long = 6         ' F: Requirements to reach this stage
Private Const QSS_COL_ONCOMPLETE As Long = 7   ' G: Effects applied when stage completes
Private Const QSS_COL_XP As Long = 8           ' H: XP reward

' ── JOURNAL TABLE COLUMN INDICES ──
Private Const JRN_COL_ID As Long = 1           ' A: EntryID
Private Const JRN_COL_DAY As Long = 2          ' B: Day recorded
Private Const JRN_COL_TIME As Long = 3         ' C: Time recorded
Private Const JRN_COL_QUESTID As Long = 4      ' D: Related quest
Private Const JRN_COL_TEXT As Long = 5         ' E: Journal text
Private Const JRN_COL_TYPE As Long = 6         ' F: Entry type (QUEST, LORE, PERSONAL)

' ── QUEST STATUS CONSTANTS ──
Public Const STATUS_INACTIVE As String = "INACTIVE"
Public Const STATUS_ACTIVE As String = "ACTIVE"
Public Const STATUS_COMPLETED As String = "COMPLETED"
Public Const STATUS_FAILED As String = "FAILED"

'===============================================================
' PUBLIC — Start a quest
'===============================================================

' Activate a quest by ID. Sets status to ACTIVE and stage to 0.
Public Sub StartQuest(questID As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then
        modUtils.DebugLog "modQuests.StartQuest: quest '" & questID & "' not found"
        Exit Sub
    End If

    ' Only start if not already active or completed
    Dim currentStatus As String
    currentStatus = UCase(modUtils.SafeStr(ws.Cells(row, QST_COL_STATUS).Value))
    If currentStatus = STATUS_ACTIVE Or currentStatus = STATUS_COMPLETED Then
        modUtils.DebugLog "modQuests.StartQuest: " & questID & " already " & currentStatus
        Exit Sub
    End If

    ws.Cells(row, QST_COL_STATUS).Value = STATUS_ACTIVE
    ws.Cells(row, QST_COL_CURSTAGE).Value = 0

    ' Load initial stage description
    Dim stageRow As Long
    stageRow = modData.GetQuestStageRow(questID, 0)
    If stageRow > 0 Then
        Dim wsQS As Worksheet
        Set wsQS = modConfig.GetSheet(modConfig.SH_QUESTSTAGES)
        If Not wsQS Is Nothing Then
            ws.Cells(row, QST_COL_STAGEDESC).Value = _
                modUtils.SafeStr(wsQS.Cells(stageRow, QSS_COL_DESC).Value)
        End If
    End If

    ' Add journal entry
    AddJournalEntry questID, "Quest started: " & GetQuestName(questID), "QUEST"

    modUtils.DebugLog "modQuests.StartQuest: activated " & questID
End Sub

'===============================================================
' PUBLIC — Advance quest to next stage
'===============================================================

' Advance a quest to the next stage. Applies stage completion
' effects and XP. Returns True if advanced successfully.
Public Function AdvanceQuest(questID As String) As Boolean
    AdvanceQuest = False

    Dim wsQ As Worksheet, wsQS As Worksheet
    Set wsQ = modConfig.GetSheet(modConfig.SH_QUESTS)
    Set wsQS = modConfig.GetSheet(modConfig.SH_QUESTSTAGES)
    If wsQ Is Nothing Or wsQS Is Nothing Then Exit Function

    Dim qRow As Long
    qRow = modData.GetQuestRow(questID)
    If qRow = 0 Then Exit Function

    ' Must be active
    If UCase(modUtils.SafeStr(wsQ.Cells(qRow, QST_COL_STATUS).Value)) <> STATUS_ACTIVE Then
        modUtils.DebugLog "modQuests.AdvanceQuest: " & questID & " is not active"
        Exit Function
    End If

    Dim currentStage As Long
    currentStage = modUtils.SafeLng(wsQ.Cells(qRow, QST_COL_CURSTAGE).Value, -1)
    Dim nextStage As Long
    nextStage = currentStage + 1

    ' Look up next stage
    Dim stageRow As Long
    stageRow = modData.GetQuestStageRow(questID, nextStage)
    If stageRow = 0 Then
        ' No more stages — auto-complete the quest
        CompleteQuest questID
        AdvanceQuest = True
        Exit Function
    End If

    ' Apply current stage completion effects
    Dim completeEffects As String
    Dim curStageRow As Long
    curStageRow = modData.GetQuestStageRow(questID, currentStage)
    If curStageRow > 0 Then
        completeEffects = modUtils.SafeStr(wsQS.Cells(curStageRow, QSS_COL_ONCOMPLETE).Value)
        If Len(completeEffects) > 0 Then
            modEffects.ProcessEffects completeEffects
        End If
    End If

    ' Award stage XP
    If curStageRow > 0 Then
        Dim xp As Long
        xp = modUtils.SafeLng(wsQS.Cells(curStageRow, QSS_COL_XP).Value, 0)
        If xp > 0 Then
            modState.AddStat modConfig.STAT_XP, xp
        End If
    End If

    ' Update quest to next stage
    wsQ.Cells(qRow, QST_COL_CURSTAGE).Value = nextStage
    wsQ.Cells(qRow, QST_COL_STAGEDESC).Value = _
        modUtils.SafeStr(wsQS.Cells(stageRow, QSS_COL_DESC).Value)

    ' Journal entry
    Dim stageName As String
    stageName = modUtils.SafeStr(wsQS.Cells(stageRow, QSS_COL_STAGENAME).Value)
    If Len(stageName) = 0 Then stageName = "Stage " & nextStage
    AddJournalEntry questID, GetQuestName(questID) & ": " & stageName, "QUEST"

    AdvanceQuest = True
    modUtils.DebugLog "modQuests.AdvanceQuest: " & questID & " -> stage " & nextStage
End Function

'===============================================================
' PUBLIC — Complete a quest
'===============================================================

' Mark a quest as COMPLETED and apply its reward effects.
Public Sub CompleteQuest(questID As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then Exit Sub

    ws.Cells(row, QST_COL_STATUS).Value = STATUS_COMPLETED
    ws.Cells(row, QST_COL_STAGEDESC).Value = "Completed"

    ' Apply quest completion rewards
    Dim rewardEff As String
    rewardEff = modUtils.SafeStr(ws.Cells(row, QST_COL_REWARD).Value)
    If Len(rewardEff) > 0 Then
        modEffects.ProcessEffects rewardEff
    End If

    AddJournalEntry questID, "Quest completed: " & GetQuestName(questID), "QUEST"
    modUtils.DebugLog "modQuests.CompleteQuest: " & questID
End Sub

'===============================================================
' PUBLIC — Fail a quest
'===============================================================

' Mark a quest as FAILED.
Public Sub FailQuest(questID As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then Exit Sub

    ws.Cells(row, QST_COL_STATUS).Value = STATUS_FAILED
    ws.Cells(row, QST_COL_STAGEDESC).Value = "Failed"

    AddJournalEntry questID, "Quest failed: " & GetQuestName(questID), "QUEST"
    modUtils.DebugLog "modQuests.FailQuest: " & questID
End Sub

'===============================================================
' PUBLIC — Quest state queries
'===============================================================

' Get quest status (INACTIVE, ACTIVE, COMPLETED, FAILED)
Public Function GetQuestStatus(questID As String) As String
    GetQuestStatus = STATUS_INACTIVE
    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then Exit Function
    GetQuestStatus = UCase(modData.ReadCellStr(modConfig.SH_QUESTS, row, QST_COL_STATUS))
    If Len(GetQuestStatus) = 0 Then GetQuestStatus = STATUS_INACTIVE
End Function

' Check if a quest is active
Public Function IsQuestActive(questID As String) As Boolean
    IsQuestActive = (GetQuestStatus(questID) = STATUS_ACTIVE)
End Function

' Check if a quest is completed
Public Function IsQuestCompleted(questID As String) As Boolean
    IsQuestCompleted = (GetQuestStatus(questID) = STATUS_COMPLETED)
End Function

' Get the current stage number of a quest
Public Function GetQuestStage(questID As String) As Long
    GetQuestStage = -1
    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then Exit Function
    GetQuestStage = modData.ReadCellLng(modConfig.SH_QUESTS, row, QST_COL_CURSTAGE)
End Function

' Get quest display name
Public Function GetQuestName(questID As String) As String
    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then Exit Function
    GetQuestName = modData.ReadCellStr(modConfig.SH_QUESTS, row, QST_COL_NAME)
End Function

' Get quest type (MAIN, SIDE, HIDDEN)
Public Function GetQuestType(questID As String) As String
    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then Exit Function
    GetQuestType = UCase(modData.ReadCellStr(modConfig.SH_QUESTS, row, QST_COL_TYPE))
End Function

' Get current stage description
Public Function GetQuestStageDesc(questID As String) As String
    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then Exit Function
    GetQuestStageDesc = modData.ReadCellStr(modConfig.SH_QUESTS, row, QST_COL_STAGEDESC)
End Function

' Get the objective text for the current stage
Public Function GetCurrentObjective(questID As String) As String
    GetCurrentObjective = ""
    Dim stage As Long
    stage = GetQuestStage(questID)
    If stage < 0 Then Exit Function

    Dim stageRow As Long
    stageRow = modData.GetQuestStageRow(questID, stage)
    If stageRow = 0 Then Exit Function

    Dim wsQS As Worksheet
    Set wsQS = modConfig.GetSheet(modConfig.SH_QUESTSTAGES)
    If wsQS Is Nothing Then Exit Function

    GetCurrentObjective = modUtils.SafeStr(wsQS.Cells(stageRow, QSS_COL_OBJECTIVE).Value)
End Function

'===============================================================
' PUBLIC — Active quest listing
'===============================================================

' Get a Collection of all active quest IDs
Public Function GetActiveQuests() As Collection
    Dim result As New Collection

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then
        Set GetActiveQuests = result
        Exit Function
    End If

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, QST_COL_ID)
        If UCase(modUtils.SafeStr(ws.Cells(r, QST_COL_STATUS).Value)) = STATUS_ACTIVE Then
            Dim qid As String
            qid = modUtils.SafeStr(ws.Cells(r, QST_COL_ID).Value)
            If Len(qid) > 0 Then result.Add qid
        End If
    Next r

    Set GetActiveQuests = result
End Function

' Get a Collection of all completed quest IDs
Public Function GetCompletedQuests() As Collection
    Dim result As New Collection

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then
        Set GetCompletedQuests = result
        Exit Function
    End If

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, QST_COL_ID)
        If UCase(modUtils.SafeStr(ws.Cells(r, QST_COL_STATUS).Value)) = STATUS_COMPLETED Then
            Dim qid As String
            qid = modUtils.SafeStr(ws.Cells(r, QST_COL_ID).Value)
            If Len(qid) > 0 Then result.Add qid
        End If
    Next r

    Set GetCompletedQuests = result
End Function

'===============================================================
' PUBLIC — Check if a quest stage's requirements are met
'===============================================================

' Check if the player meets the requirements for the current stage.
' Used by the scene engine to auto-advance quests at checkpoints.
Public Function CanAdvanceQuest(questID As String) As Boolean
    CanAdvanceQuest = False

    If Not IsQuestActive(questID) Then Exit Function

    Dim stage As Long
    stage = GetQuestStage(questID)
    If stage < 0 Then Exit Function

    Dim stageRow As Long
    stageRow = modData.GetQuestStageRow(questID, stage)
    If stageRow = 0 Then Exit Function

    Dim wsQS As Worksheet
    Set wsQS = modConfig.GetSheet(modConfig.SH_QUESTSTAGES)
    If wsQS Is Nothing Then Exit Function

    Dim reqs As String
    reqs = modUtils.SafeStr(wsQS.Cells(stageRow, QSS_COL_REQS).Value)

    ' If no requirements, can always advance
    If Len(reqs) = 0 Then
        CanAdvanceQuest = True
    Else
        CanAdvanceQuest = modRequirements.CheckRequirements(reqs)
    End If
End Function

'===============================================================
' PUBLIC — Journal system
'===============================================================

' Add a journal entry to the journal table
Public Sub AddJournalEntry(relatedQuestID As String, text As String, entryType As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Sub

    ' Find next empty row
    Dim nextRow As Long
    nextRow = modUtils.GetLastRow(ws, JRN_COL_ID) + 1

    ' Generate entry ID
    Dim entryID As String
    entryID = "JRN_" & Format(nextRow - 1, "000")

    ws.Cells(nextRow, JRN_COL_ID).Value = entryID
    ws.Cells(nextRow, JRN_COL_DAY).Value = modState.GetCurrentDay()
    ws.Cells(nextRow, JRN_COL_TIME).Value = modState.GetTimeOfDay()
    ws.Cells(nextRow, JRN_COL_QUESTID).Value = relatedQuestID
    ws.Cells(nextRow, JRN_COL_TEXT).Value = text
    ws.Cells(nextRow, JRN_COL_TYPE).Value = entryType

    modUtils.DebugLog "modQuests.AddJournalEntry: " & entryID & " - " & text
End Sub

' Get the most recent N journal entries as formatted text
Public Function GetRecentJournal(Optional maxEntries As Long = 5) As String
    GetRecentJournal = ""

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, JRN_COL_ID)
    If lastRow < 2 Then
        GetRecentJournal = "(No journal entries)"
        Exit Function
    End If

    Dim startRow As Long
    startRow = lastRow - maxEntries + 1
    If startRow < 2 Then startRow = 2

    Dim result As String
    Dim r As Long
    For r = lastRow To startRow Step -1
        Dim day As Long
        day = modUtils.SafeLng(ws.Cells(r, JRN_COL_DAY).Value, 0)
        Dim timeSlot As String
        timeSlot = modUtils.SafeStr(ws.Cells(r, JRN_COL_TIME).Value)
        Dim entryText As String
        entryText = modUtils.SafeStr(ws.Cells(r, JRN_COL_TEXT).Value)

        result = result & "Day " & day & " (" & timeSlot & "): " & entryText & vbLf
    Next r

    GetRecentJournal = result
End Function
