Attribute VB_Name = "modData"
'===============================================================
' modData — Data Caches & Lookup Functions
' Damned Moon VBA RPG Engine
'===============================================================
' Builds dictionary caches from workbook tables during init.
' All row lookups go through here — no more wandering row by row.
'===============================================================

Option Explicit

' ── CACHE DICTIONARIES ──
' Each maps an ID string -> row number (Long) on its respective sheet
Public gSceneRows As Object      ' Scripting.Dictionary: SceneID -> row
Public gStatRows As Object       ' Scripting.Dictionary: StatName -> row
Public gFlagRows As Object       ' Scripting.Dictionary: FlagName -> row
Public gItemRows As Object       ' Scripting.Dictionary: ItemID -> row
Public gQuestRows As Object      ' Scripting.Dictionary: QuestID -> row
Public gEnemyRows As Object      ' Scripting.Dictionary: EnemyID -> row
Public gMapNodeRows As Object    ' Scripting.Dictionary: NodeID -> row
Public gMapLinkRows As Object    ' Scripting.Dictionary: "FromID|ToID" -> row
Public gNPCRows As Object        ' Scripting.Dictionary: NPCID -> row
Public gEncounterRows As Object  ' Scripting.Dictionary: EncounterID -> row
Public gQuestStageRows As Object ' Scripting.Dictionary: "QuestID|StageNum" -> row
Public gJobRows As Object        ' Scripting.Dictionary: JobID -> row
Public gJournalRows As Object    ' Scripting.Dictionary: EntryID -> row
Public gMoonPhaseRows As Object  ' Scripting.Dictionary: PhaseID -> row

' ── CACHE STATE ──
Private mCachesBuilt As Boolean

'===============================================================
' BUILD ALL CACHES — call once during initialization
'===============================================================
Public Sub BuildCaches()
    modUtils.DebugLog "modData.BuildCaches: starting"

    Set gSceneRows = CreateObject("Scripting.Dictionary")
    Set gStatRows = CreateObject("Scripting.Dictionary")
    Set gFlagRows = CreateObject("Scripting.Dictionary")
    Set gItemRows = CreateObject("Scripting.Dictionary")
    Set gQuestRows = CreateObject("Scripting.Dictionary")
    Set gEnemyRows = CreateObject("Scripting.Dictionary")
    Set gMapNodeRows = CreateObject("Scripting.Dictionary")
    Set gMapLinkRows = CreateObject("Scripting.Dictionary")
    Set gNPCRows = CreateObject("Scripting.Dictionary")
    Set gEncounterRows = CreateObject("Scripting.Dictionary")
    Set gQuestStageRows = CreateObject("Scripting.Dictionary")
    Set gJobRows = CreateObject("Scripting.Dictionary")
    Set gJournalRows = CreateObject("Scripting.Dictionary")
    Set gMoonPhaseRows = CreateObject("Scripting.Dictionary")

    ' Build each cache from its sheet (col 1 = ID, except where noted)
    BuildSingleCache gSceneRows, modConfig.SH_SCENES, 1
    BuildSingleCache gStatRows, modConfig.SH_STATS, 1
    BuildSingleCache gFlagRows, modConfig.SH_FLAGS, 1
    BuildSingleCache gItemRows, modConfig.SH_ITEMS, 1
    BuildSingleCache gQuestRows, modConfig.SH_QUESTS, 1
    BuildSingleCache gEnemyRows, modConfig.SH_ENEMIES, 1
    BuildSingleCache gMapNodeRows, modConfig.SH_MAPNODES, 1
    BuildSingleCache gNPCRows, modConfig.SH_NPCS, 1
    BuildSingleCache gEncounterRows, modConfig.SH_ENCOUNTERS, 1
    BuildSingleCache gJobRows, modConfig.SH_JOBS, 1
    BuildSingleCache gJournalRows, modConfig.SH_JOURNAL, 1
    BuildSingleCache gMoonPhaseRows, modConfig.SH_MOON, 1

    ' Map links use composite key: "FromID|ToID"
    BuildMapLinkCache

    ' Quest stages use composite key: "QuestID|StageNum"
    BuildQuestStageCache

    mCachesBuilt = True
    modUtils.DebugLog "modData.BuildCaches: complete"
End Sub

'===============================================================
' CACHE STATUS
'===============================================================
Public Function AreCachesBuilt() As Boolean
    AreCachesBuilt = mCachesBuilt
End Function

'===============================================================
' SINGLE-KEY CACHE BUILDER
'===============================================================
Private Sub BuildSingleCache(dict As Object, sheetName As String, keyCol As Long)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(sheetName)
    If ws Is Nothing Then
        modUtils.DebugLog "modData.BuildSingleCache: sheet '" & sheetName & "' not found, skipping"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, keyCol)

    Dim r As Long
    For r = 2 To lastRow
        Dim key As String
        key = modUtils.SafeStr(ws.Cells(r, keyCol).Value)
        If key <> "" Then
            If Not dict.Exists(key) Then
                dict.Add key, r
            End If
        End If
    Next r

    modUtils.DebugLog "  cached " & dict.Count & " rows from " & sheetName
End Sub

'===============================================================
' MAP LINK CACHE (composite key)
'===============================================================
Private Sub BuildMapLinkCache()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_MAPLINKS)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, 1)

    Dim r As Long
    For r = 2 To lastRow
        Dim fromID As String
        Dim toID As String
        fromID = modUtils.SafeStr(ws.Cells(r, 1).Value)
        toID = modUtils.SafeStr(ws.Cells(r, 2).Value)
        If fromID <> "" And toID <> "" Then
            Dim linkKey As String
            linkKey = fromID & "|" & toID
            If Not gMapLinkRows.Exists(linkKey) Then
                gMapLinkRows.Add linkKey, r
            End If
        End If
    Next r

    modUtils.DebugLog "  cached " & gMapLinkRows.Count & " map links"
End Sub

'===============================================================
' QUEST STAGE CACHE (composite key)
'===============================================================
Private Sub BuildQuestStageCache()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTSTAGES)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, 1)

    Dim r As Long
    For r = 2 To lastRow
        Dim questID As String
        Dim stageNum As String
        questID = modUtils.SafeStr(ws.Cells(r, 1).Value)
        stageNum = modUtils.SafeStr(ws.Cells(r, 2).Value)
        If questID <> "" And stageNum <> "" Then
            Dim stageKey As String
            stageKey = questID & "|" & stageNum
            If Not gQuestStageRows.Exists(stageKey) Then
                gQuestStageRows.Add stageKey, r
            End If
        End If
    Next r

    modUtils.DebugLog "  cached " & gQuestStageRows.Count & " quest stages"
End Sub

'===============================================================
' ROW LOOKUP FUNCTIONS
'===============================================================

' Get the row number for a SceneID. Returns 0 if not found.
Public Function GetSceneRow(sceneID As String) As Long
    If gSceneRows Is Nothing Then
        GetSceneRow = 0
    ElseIf gSceneRows.Exists(sceneID) Then
        GetSceneRow = gSceneRows(sceneID)
    Else
        GetSceneRow = 0
    End If
End Function

' Get the row number for a stat name. Returns 0 if not found.
Public Function GetStatRow(statName As String) As Long
    If gStatRows Is Nothing Then
        GetStatRow = 0
    ElseIf gStatRows.Exists(statName) Then
        GetStatRow = gStatRows(statName)
    Else
        GetStatRow = 0
    End If
End Function

' Get the row number for a flag name. Returns 0 if not found.
Public Function GetFlagRow(flagName As String) As Long
    If gFlagRows Is Nothing Then
        GetFlagRow = 0
    ElseIf gFlagRows.Exists(flagName) Then
        GetFlagRow = gFlagRows(flagName)
    Else
        GetFlagRow = 0
    End If
End Function

' Get the row number for an item ID. Returns 0 if not found.
Public Function GetItemRow(itemID As String) As Long
    If gItemRows Is Nothing Then
        GetItemRow = 0
    ElseIf gItemRows.Exists(itemID) Then
        GetItemRow = gItemRows(itemID)
    Else
        GetItemRow = 0
    End If
End Function

' Get the row number for a quest ID. Returns 0 if not found.
Public Function GetQuestRow(questID As String) As Long
    If gQuestRows Is Nothing Then
        GetQuestRow = 0
    ElseIf gQuestRows.Exists(questID) Then
        GetQuestRow = gQuestRows(questID)
    Else
        GetQuestRow = 0
    End If
End Function

' Get the row number for an enemy ID. Returns 0 if not found.
Public Function GetEnemyRow(enemyID As String) As Long
    If gEnemyRows Is Nothing Then
        GetEnemyRow = 0
    ElseIf gEnemyRows.Exists(enemyID) Then
        GetEnemyRow = gEnemyRows(enemyID)
    Else
        GetEnemyRow = 0
    End If
End Function

' Get the row number for a map node ID. Returns 0 if not found.
Public Function GetMapNodeRow(nodeID As String) As Long
    If gMapNodeRows Is Nothing Then
        GetMapNodeRow = 0
    ElseIf gMapNodeRows.Exists(nodeID) Then
        GetMapNodeRow = gMapNodeRows(nodeID)
    Else
        GetMapNodeRow = 0
    End If
End Function

' Get the row number for a map link. Returns 0 if not found.
Public Function GetMapLinkRow(fromID As String, toID As String) As Long
    If gMapLinkRows Is Nothing Then
        GetMapLinkRow = 0
        Exit Function
    End If
    Dim linkKey As String
    linkKey = fromID & "|" & toID
    If gMapLinkRows.Exists(linkKey) Then
        GetMapLinkRow = gMapLinkRows(linkKey)
    Else
        GetMapLinkRow = 0
    End If
End Function

' Get the row number for an NPC ID. Returns 0 if not found.
Public Function GetNPCRow(npcID As String) As Long
    If gNPCRows Is Nothing Then
        GetNPCRow = 0
    ElseIf gNPCRows.Exists(npcID) Then
        GetNPCRow = gNPCRows(npcID)
    Else
        GetNPCRow = 0
    End If
End Function

' Get the row number for an encounter ID. Returns 0 if not found.
Public Function GetEncounterRow(encounterID As String) As Long
    If gEncounterRows Is Nothing Then
        GetEncounterRow = 0
    ElseIf gEncounterRows.Exists(encounterID) Then
        GetEncounterRow = gEncounterRows(encounterID)
    Else
        GetEncounterRow = 0
    End If
End Function

' Get the row number for a quest stage. Returns 0 if not found.
Public Function GetQuestStageRow(questID As String, stageNum As Long) As Long
    If gQuestStageRows Is Nothing Then
        GetQuestStageRow = 0
        Exit Function
    End If
    Dim stageKey As String
    stageKey = questID & "|" & CStr(stageNum)
    If gQuestStageRows.Exists(stageKey) Then
        GetQuestStageRow = gQuestStageRows(stageKey)
    Else
        GetQuestStageRow = 0
    End If
End Function

' Get the row number for a job ID. Returns 0 if not found.
Public Function GetJobRow(jobID As String) As Long
    If gJobRows Is Nothing Then
        GetJobRow = 0
    ElseIf gJobRows.Exists(jobID) Then
        GetJobRow = gJobRows(jobID)
    Else
        GetJobRow = 0
    End If
End Function

' Get the row number for a journal entry ID. Returns 0 if not found.
Public Function GetJournalRow(entryID As String) As Long
    If gJournalRows Is Nothing Then
        GetJournalRow = 0
    ElseIf gJournalRows.Exists(entryID) Then
        GetJournalRow = gJournalRows(entryID)
    Else
        GetJournalRow = 0
    End If
End Function

'===============================================================
' EXISTENCE CHECKS
'===============================================================

Public Function SceneExists(sceneID As String) As Boolean
    SceneExists = (GetSceneRow(sceneID) > 0)
End Function

Public Function ItemExists(itemID As String) As Boolean
    ItemExists = (GetItemRow(itemID) > 0)
End Function

Public Function QuestExists(questID As String) As Boolean
    QuestExists = (GetQuestRow(questID) > 0)
End Function

Public Function EnemyExists(enemyID As String) As Boolean
    EnemyExists = (GetEnemyRow(enemyID) > 0)
End Function

Public Function MapNodeExists(nodeID As String) As Boolean
    MapNodeExists = (GetMapNodeRow(nodeID) > 0)
End Function

Public Function MapLinkExists(fromID As String, toID As String) As Boolean
    MapLinkExists = (GetMapLinkRow(fromID, toID) > 0)
End Function

'===============================================================
' CELL VALUE READERS (cached row + column)
'===============================================================

' Read a cell value from a cached sheet+row by column number
Public Function ReadCell(sheetName As String, row As Long, col As Long) As Variant
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(sheetName)
    If ws Is Nothing Or row = 0 Then
        ReadCell = ""
        Exit Function
    End If
    ReadCell = ws.Cells(row, col).Value
End Function

' Read a cell value as String
Public Function ReadCellStr(sheetName As String, row As Long, col As Long) As String
    ReadCellStr = modUtils.SafeStr(ReadCell(sheetName, row, col))
End Function

' Read a cell value as Long
Public Function ReadCellLng(sheetName As String, row As Long, col As Long) As Long
    ReadCellLng = modUtils.SafeLng(ReadCell(sheetName, row, col))
End Function

'===============================================================
' INVALIDATION — call if table data changes at runtime
'===============================================================
Public Sub InvalidateCaches()
    mCachesBuilt = False
    Set gSceneRows = Nothing
    Set gStatRows = Nothing
    Set gFlagRows = Nothing
    Set gItemRows = Nothing
    Set gQuestRows = Nothing
    Set gEnemyRows = Nothing
    Set gMapNodeRows = Nothing
    Set gMapLinkRows = Nothing
    Set gNPCRows = Nothing
    Set gEncounterRows = Nothing
    Set gQuestStageRows = Nothing
    Set gJobRows = Nothing
    Set gJournalRows = Nothing
    Set gMoonPhaseRows = Nothing
    modUtils.DebugLog "modData.InvalidateCaches: all caches cleared"
End Sub
