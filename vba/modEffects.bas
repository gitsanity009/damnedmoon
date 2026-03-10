Attribute VB_Name = "modEffects"
'===============================================================
' modEffects — Effect Processing Engine
' Damned Moon VBA RPG Engine — Phase 2
'===============================================================
' Applies effect strings from scenes and choices. Effects modify
' game state: stats, flags, items, quests, time, and triggers.
'
' Effect syntax (pipe-delimited for multiple):
'   STAT:HEALTH+5            — add to stat
'   STAT:RAGE-10             — subtract from stat
'   STAT:HUNGER=50           — set stat to value
'   FLAG_SET:MetHunter        — set flag to TRUE
'   FLAG_CLEAR:MetHunter      — set flag to FALSE
'   FLAG_TOGGLE:HiddenPath    — toggle flag
'   ITEM_ADD:ITM_KNIFE        — add item to inventory
'   ITEM_REMOVE:ITM_BANDAGE   — remove item from inventory
'   QUEST_START:QST_CURE      — activate a quest
'   QUEST_ADVANCE:QST_CURE    — advance quest to next stage
'   TIME_ADVANCE:60           — advance time by N minutes
'   DAY_ADVANCE:1             — advance day by N
'   SCENE_JUMP:SCN_BLACKOUT   — force jump to another scene
'===============================================================

Option Explicit

' ── SCENE JUMP FLAG ──
' If an effect triggers a scene jump, store it here for the caller
Private mPendingJump As String

'===============================================================
' PUBLIC — Process a full effect string
'===============================================================

' Process all effects in a pipe-delimited string.
' Returns the pending scene jump ID if any (or "" for none).
Public Function ProcessEffects(effectStr As String) As String
    mPendingJump = ""

    If Len(Trim(effectStr)) = 0 Then
        ProcessEffects = ""
        Exit Function
    End If

    Dim effects As Variant
    effects = modUtils.SplitTrimmed(effectStr, modConfig.EFFECT_DELIM)

    Dim i As Long
    For i = LBound(effects) To UBound(effects)
        Dim ef As String
        ef = Trim(CStr(effects(i)))
        If Len(ef) > 0 Then
            ProcessSingleEffect ef
        End If
    Next i

    ProcessEffects = mPendingJump
End Function

'===============================================================
' PRIVATE — Single effect dispatcher
'===============================================================
Private Sub ProcessSingleEffect(ef As String)
    If modUtils.StartsWith(ef, "STAT:") Then
        ProcessStatEffect modUtils.StripPrefix(ef, "STAT:")

    ElseIf modUtils.StartsWith(ef, "FLAG_SET:") Then
        modState.SetFlag modUtils.StripPrefix(ef, "FLAG_SET:"), True

    ElseIf modUtils.StartsWith(ef, "FLAG_CLEAR:") Then
        modState.SetFlag modUtils.StripPrefix(ef, "FLAG_CLEAR:"), False

    ElseIf modUtils.StartsWith(ef, "FLAG_TOGGLE:") Then
        modState.ToggleFlag modUtils.StripPrefix(ef, "FLAG_TOGGLE:")

    ElseIf modUtils.StartsWith(ef, "ITEM_ADD:") Then
        AddItem modUtils.StripPrefix(ef, "ITEM_ADD:")

    ElseIf modUtils.StartsWith(ef, "ITEM_REMOVE:") Then
        RemoveItem modUtils.StripPrefix(ef, "ITEM_REMOVE:")

    ElseIf modUtils.StartsWith(ef, "QUEST_START:") Then
        StartQuest modUtils.StripPrefix(ef, "QUEST_START:")

    ElseIf modUtils.StartsWith(ef, "QUEST_ADVANCE:") Then
        AdvanceQuest modUtils.StripPrefix(ef, "QUEST_ADVANCE:")

    ElseIf modUtils.StartsWith(ef, "TIME_ADVANCE:") Then
        Dim mins As Long
        mins = modUtils.SafeLng(modUtils.StripPrefix(ef, "TIME_ADVANCE:"), 0)
        If mins > 0 Then modState.AdvanceTime mins

    ElseIf modUtils.StartsWith(ef, "DAY_ADVANCE:") Then
        Dim days As Long
        days = modUtils.SafeLng(modUtils.StripPrefix(ef, "DAY_ADVANCE:"), 0)
        Dim d As Long
        For d = 1 To days
            modState.AdvanceDay
        Next d

    ElseIf modUtils.StartsWith(ef, "SCENE_JUMP:") Then
        mPendingJump = modUtils.StripPrefix(ef, "SCENE_JUMP:")
        modUtils.DebugLog "modEffects: queued scene jump to " & mPendingJump

    Else
        modUtils.DebugLog "modEffects: unknown effect '" & ef & "', skipping"
    End If
End Sub

'===============================================================
' STAT EFFECT: HEALTH+5, RAGE-10, HUNGER=50
'===============================================================
Private Sub ProcessStatEffect(statStr As String)
    Dim opPos As Long
    opPos = modUtils.FindOperatorPos(statStr)
    If opPos = 0 Then
        modUtils.DebugLog "modEffects.ProcessStatEffect: no operator in '" & statStr & "'"
        Exit Sub
    End If

    Dim statName As String
    statName = Left(statStr, opPos - 1)

    Dim op As String
    op = Mid(statStr, opPos, 1)

    Dim val As Long
    val = modUtils.SafeLng(Mid(statStr, opPos + 1), 0)

    Select Case op
        Case "+": modState.AddStat statName, val
        Case "-": modState.AddStat statName, -val
        Case "=": modState.SetStat statName, val
    End Select

    ' Check RAGE 100 blackout trigger
    If statName = modConfig.STAT_RAGE Then
        If modState.GetStat(modConfig.STAT_RAGE) >= 100 Then
            mPendingJump = "SCN_BLACKOUT"
            modUtils.DebugLog "modEffects: RAGE>=100, queued blackout jump"
        End If
    End If
End Sub

'===============================================================
' INVENTORY — Add / Remove items
'===============================================================
Private Sub AddItem(itemID As String)
    Dim wsInv As Worksheet, wsItems As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    Set wsItems = modConfig.GetSheet(modConfig.SH_ITEMS)
    If wsInv Is Nothing Then Exit Sub

    ' Look up item name from DB
    Dim itemName As String
    Dim itemRow As Long
    itemRow = modData.GetItemRow(itemID)
    If itemRow > 0 And Not wsItems Is Nothing Then
        itemName = modUtils.SafeStr(wsItems.Cells(itemRow, 2).Value)
    Else
        itemName = itemID  ' fallback to ID if not in DB
    End If

    ' Check if already in inventory (stack)
    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, 1)
        If modUtils.SafeStr(wsInv.Cells(r, 2).Value) = itemID Then
            wsInv.Cells(r, 4).Value = modUtils.SafeLng(wsInv.Cells(r, 4).Value, 0) + 1
            modUtils.DebugLog "modEffects.AddItem: stacked " & itemID
            Exit Sub
        End If
    Next r

    ' Find empty slot
    For r = 2 To modUtils.GetLastRow(wsInv, 1)
        Dim slotID As String
        slotID = modUtils.SafeStr(wsInv.Cells(r, 2).Value)
        Dim slotQty As Long
        slotQty = modUtils.SafeLng(wsInv.Cells(r, 4).Value, 0)
        If slotID = "" And slotQty = 0 Then
            wsInv.Cells(r, 2).Value = itemID
            wsInv.Cells(r, 3).Value = itemName
            wsInv.Cells(r, 4).Value = 1
            modUtils.DebugLog "modEffects.AddItem: added " & itemID
            Exit Sub
        End If
    Next r

    modUtils.DebugLog "modEffects.AddItem: no empty slot for " & itemID
End Sub

Private Sub RemoveItem(itemID As String)
    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Sub

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, 1)
        If modUtils.SafeStr(wsInv.Cells(r, 2).Value) = itemID Then
            Dim qty As Long
            qty = modUtils.SafeLng(wsInv.Cells(r, 4).Value, 0) - 1
            If qty <= 0 Then
                wsInv.Cells(r, 2).Value = ""
                wsInv.Cells(r, 3).Value = ""
                wsInv.Cells(r, 4).Value = 0
            Else
                wsInv.Cells(r, 4).Value = qty
            End If
            modUtils.DebugLog "modEffects.RemoveItem: " & itemID & " (qty=" & qty & ")"
            Exit Sub
        End If
    Next r
End Sub

'===============================================================
' QUEST — Start / Advance
'===============================================================
Private Sub StartQuest(questID As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetQuestRow(questID)
    If row = 0 Then
        modUtils.DebugLog "modEffects.StartQuest: quest '" & questID & "' not found"
        Exit Sub
    End If

    ws.Cells(row, 4).Value = "ACTIVE"
    ws.Cells(row, 6).Value = 0
    modUtils.DebugLog "modEffects.StartQuest: activated " & questID
End Sub

Private Sub AdvanceQuest(questID As String)
    Dim wsQ As Worksheet, wsQS As Worksheet
    Set wsQ = modConfig.GetSheet(modConfig.SH_QUESTS)
    Set wsQS = modConfig.GetSheet(modConfig.SH_QUESTSTAGES)
    If wsQ Is Nothing Or wsQS Is Nothing Then Exit Sub

    Dim qRow As Long
    qRow = modData.GetQuestRow(questID)
    If qRow = 0 Then Exit Sub

    Dim currentStage As Long
    currentStage = modUtils.SafeLng(wsQ.Cells(qRow, 6).Value, -1)

    Dim nextStage As Long
    nextStage = currentStage + 1

    ' Look up next stage row
    Dim stageRow As Long
    stageRow = modData.GetQuestStageRow(questID, nextStage)
    If stageRow = 0 Then
        modUtils.DebugLog "modEffects.AdvanceQuest: no stage " & nextStage & " for " & questID
        Exit Sub
    End If

    ' Update quest
    wsQ.Cells(qRow, 6).Value = nextStage
    wsQ.Cells(qRow, 5).Value = modUtils.SafeStr(wsQS.Cells(stageRow, 4).Value) ' stage description

    ' Award stage XP
    Dim xp As Long
    xp = modUtils.SafeLng(wsQS.Cells(stageRow, 8).Value, 0)
    If xp > 0 Then
        modState.AddStat modConfig.STAT_XP, xp
    End If

    modUtils.DebugLog "modEffects.AdvanceQuest: " & questID & " -> stage " & nextStage
End Sub
