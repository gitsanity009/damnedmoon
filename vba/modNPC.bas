Attribute VB_Name = "modNPC"
'===============================================================
' modNPC — NPC Interaction System
' Damned Moon VBA RPG Engine — Phase 4
'===============================================================
' Manages NPC data, disposition (affinity), dialogue selection,
' location tracking, gift-giving, and suspicion. NPCs remember
' player interactions through affinity scores and flags.
'
' Data source: tbl_NPCs
'   A: NPCID        B: Name         C: Title/Role
'   D: Description  E: HomeLocation F: CurrentLocation
'   G: Affinity     H: Suspicion    I: IsAlive
'   J: DialogueDefault  K: DialogueFriendly  L: DialogueHostile
'   M: DialogueSpecial  N: SpecialReqs
'   O: GiftEffects  P: Faction      Q: Schedule
'
' Affinity scale: -100 (hostile) to +100 (devoted)
' Suspicion scale: 0 (unsuspecting) to 100 (knows you're a wolf)
'===============================================================

Option Explicit

' ── NPC TABLE COLUMN INDICES ──
Private Const NPC_COL_ID As Long = 1            ' A: NPCID
Private Const NPC_COL_NAME As Long = 2          ' B: Display name
Private Const NPC_COL_TITLE As Long = 3         ' C: Title / role
Private Const NPC_COL_DESC As Long = 4          ' D: Description
Private Const NPC_COL_HOME As Long = 5          ' E: Home location (node ID)
Private Const NPC_COL_LOCATION As Long = 6      ' F: Current location (node ID)
Private Const NPC_COL_AFFINITY As Long = 7      ' G: Affinity score (-100 to +100)
Private Const NPC_COL_SUSPICION As Long = 8     ' H: Suspicion score (0 to 100)
Private Const NPC_COL_ALIVE As Long = 9         ' I: IsAlive flag
Private Const NPC_COL_DLG_DEFAULT As Long = 10  ' J: Default dialogue text
Private Const NPC_COL_DLG_FRIENDLY As Long = 11 ' K: Friendly dialogue (affinity > 30)
Private Const NPC_COL_DLG_HOSTILE As Long = 12  ' L: Hostile dialogue (affinity < -30)
Private Const NPC_COL_DLG_SPECIAL As Long = 13  ' M: Special dialogue (conditional)
Private Const NPC_COL_SPECIAL_REQS As Long = 14 ' N: Requirements for special dialogue
Private Const NPC_COL_GIFT_EFFECTS As Long = 15 ' O: Effect string applied when gifted items
Private Const NPC_COL_FACTION As Long = 16      ' P: Faction (VILLAGER, HUNTER, PACK, CHURCH, LONER)
Private Const NPC_COL_SCHEDULE As Long = 17     ' Q: Schedule (pipe-delim: TIME:LOCATION pairs)

' ── AFFINITY THRESHOLDS ──
Private Const AFFINITY_FRIENDLY As Long = 30
Private Const AFFINITY_HOSTILE As Long = -30
Private Const AFFINITY_MIN As Long = -100
Private Const AFFINITY_MAX As Long = 100

' ── SUSPICION THRESHOLDS ──
Private Const SUSPICION_WARY As Long = 30
Private Const SUSPICION_ALERT As Long = 60
Private Const SUSPICION_KNOWS As Long = 90

'===============================================================
' PUBLIC — NPC Data Lookups
'===============================================================

' Get NPC display name
Public Function GetNPCName(npcID As String) As String
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function
    GetNPCName = modData.ReadCellStr(modConfig.SH_NPCS, row, NPC_COL_NAME)
End Function

' Get NPC title/role
Public Function GetNPCTitle(npcID As String) As String
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function
    GetNPCTitle = modData.ReadCellStr(modConfig.SH_NPCS, row, NPC_COL_TITLE)
End Function

' Get NPC description
Public Function GetNPCDescription(npcID As String) As String
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function
    GetNPCDescription = modData.ReadCellStr(modConfig.SH_NPCS, row, NPC_COL_DESC)
End Function

' Get NPC faction
Public Function GetNPCFaction(npcID As String) As String
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function
    GetNPCFaction = UCase(modData.ReadCellStr(modConfig.SH_NPCS, row, NPC_COL_FACTION))
End Function

' Check if NPC is alive
Public Function IsNPCAlive(npcID As String) As Boolean
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then
        IsNPCAlive = False
        Exit Function
    End If
    IsNPCAlive = modUtils.SafeBool(modData.ReadCell(modConfig.SH_NPCS, row, NPC_COL_ALIVE), True)
End Function

'===============================================================
' PUBLIC — Affinity System
'===============================================================

' Get NPC's affinity toward the player (-100 to +100)
Public Function GetAffinity(npcID As String) As Long
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function
    GetAffinity = modData.ReadCellLng(modConfig.SH_NPCS, row, NPC_COL_AFFINITY)
End Function

' Set NPC affinity (clamped to -100..+100)
Public Sub SetAffinity(npcID As String, val As Long)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
    If ws Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Sub

    ws.Cells(row, NPC_COL_AFFINITY).Value = modUtils.Clamp(val, AFFINITY_MIN, AFFINITY_MAX)
    modUtils.DebugLog "modNPC.SetAffinity: " & npcID & " = " & val
End Sub

' Adjust NPC affinity by a delta (clamped)
Public Sub AdjustAffinity(npcID As String, delta As Long)
    Dim current As Long
    current = GetAffinity(npcID)
    SetAffinity npcID, current + delta
End Sub

' Check if NPC is friendly (affinity >= threshold)
Public Function IsFriendly(npcID As String) As Boolean
    IsFriendly = (GetAffinity(npcID) >= AFFINITY_FRIENDLY)
End Function

' Check if NPC is hostile (affinity <= threshold)
Public Function IsHostile(npcID As String) As Boolean
    IsHostile = (GetAffinity(npcID) <= AFFINITY_HOSTILE)
End Function

'===============================================================
' PUBLIC — Suspicion System
'===============================================================

' Get NPC's suspicion level (0 to 100)
Public Function GetSuspicion(npcID As String) As Long
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function
    GetSuspicion = modData.ReadCellLng(modConfig.SH_NPCS, row, NPC_COL_SUSPICION)
End Function

' Set NPC suspicion (clamped 0-100)
Public Sub SetSuspicion(npcID As String, val As Long)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
    If ws Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Sub

    ws.Cells(row, NPC_COL_SUSPICION).Value = modUtils.Clamp(val, 0, 100)
    modUtils.DebugLog "modNPC.SetSuspicion: " & npcID & " = " & val
End Sub

' Adjust NPC suspicion by a delta
Public Sub AdjustSuspicion(npcID As String, delta As Long)
    Dim current As Long
    current = GetSuspicion(npcID)
    SetSuspicion npcID, current + delta
End Sub

' Check if NPC is wary (suspicion >= 30)
Public Function IsWary(npcID As String) As Boolean
    IsWary = (GetSuspicion(npcID) >= SUSPICION_WARY)
End Function

' Check if NPC is on alert (suspicion >= 60)
Public Function IsAlert(npcID As String) As Boolean
    IsAlert = (GetSuspicion(npcID) >= SUSPICION_ALERT)
End Function

' Check if NPC knows (suspicion >= 90)
Public Function NPCKnows(npcID As String) As Boolean
    NPCKnows = (GetSuspicion(npcID) >= SUSPICION_KNOWS)
End Function

'===============================================================
' PUBLIC — NPC Location & Schedule
'===============================================================

' Get NPC's current location (node ID)
Public Function GetNPCLocation(npcID As String) As String
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function
    GetNPCLocation = modData.ReadCellStr(modConfig.SH_NPCS, row, NPC_COL_LOCATION)
End Function

' Get NPC's home location
Public Function GetNPCHome(npcID As String) As String
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function
    GetNPCHome = modData.ReadCellStr(modConfig.SH_NPCS, row, NPC_COL_HOME)
End Function

' Move an NPC to a new location
Public Sub MoveNPC(npcID As String, newLocation As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
    If ws Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Sub

    ws.Cells(row, NPC_COL_LOCATION).Value = newLocation
    modUtils.DebugLog "modNPC.MoveNPC: " & npcID & " -> " & newLocation
End Sub

' Update all NPC locations based on their schedules and current time.
' Schedule format: "MORNING:NODE_INN|AFTERNOON:NODE_STORE|NIGHT:NODE_HOME"
Public Sub UpdateNPCSchedules()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
    If ws Is Nothing Then Exit Sub

    Dim currentTime As String
    currentTime = UCase(modState.GetTimeOfDay())

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, NPC_COL_ID)
        Dim npcID As String
        npcID = modUtils.SafeStr(ws.Cells(r, NPC_COL_ID).Value)
        If Len(npcID) = 0 Then GoTo NextNPC

        ' Skip dead NPCs
        If Not modUtils.SafeBool(ws.Cells(r, NPC_COL_ALIVE).Value, True) Then GoTo NextNPC

        Dim schedule As String
        schedule = modUtils.SafeStr(ws.Cells(r, NPC_COL_SCHEDULE).Value)
        If Len(schedule) = 0 Then GoTo NextNPC

        ' Parse schedule entries
        Dim entries As Variant
        entries = modUtils.SplitTrimmed(schedule, "|")

        Dim i As Long
        For i = LBound(entries) To UBound(entries)
            Dim entry As String
            entry = Trim(CStr(entries(i)))
            If InStr(entry, ":") > 0 Then
                Dim timePart As String
                Dim locPart As String
                timePart = UCase(Trim(Left(entry, InStr(entry, ":") - 1)))
                locPart = Trim(Mid(entry, InStr(entry, ":") + 1))

                If timePart = currentTime Then
                    ws.Cells(r, NPC_COL_LOCATION).Value = locPart
                    Exit For
                End If
            End If
        Next i
NextNPC:
    Next r

    modUtils.DebugLog "modNPC.UpdateNPCSchedules: updated for time " & currentTime
End Sub

' Get all NPCs at a given location. Returns a Collection of NPC IDs.
Public Function GetNPCsAtLocation(nodeID As String) As Collection
    Dim result As New Collection

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
    If ws Is Nothing Then
        Set GetNPCsAtLocation = result
        Exit Function
    End If

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, NPC_COL_ID)
        Dim npcID As String
        npcID = modUtils.SafeStr(ws.Cells(r, NPC_COL_ID).Value)
        If Len(npcID) = 0 Then GoTo NextLoc

        ' Must be alive
        If Not modUtils.SafeBool(ws.Cells(r, NPC_COL_ALIVE).Value, True) Then GoTo NextLoc

        ' Check location match
        Dim npcLoc As String
        npcLoc = UCase(modUtils.SafeStr(ws.Cells(r, NPC_COL_LOCATION).Value))
        If npcLoc = UCase(nodeID) Then
            result.Add npcID
        End If
NextLoc:
    Next r

    Set GetNPCsAtLocation = result
End Function

'===============================================================
' PUBLIC — Dialogue
'===============================================================

' Get the appropriate dialogue text for an NPC based on affinity,
' suspicion, and special requirements.
Public Function GetDialogue(npcID As String) As String
    GetDialogue = ""

    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Function

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
    If ws Is Nothing Then Exit Function

    ' Check special dialogue first (highest priority)
    Dim specialReqs As String
    specialReqs = modUtils.SafeStr(ws.Cells(row, NPC_COL_SPECIAL_REQS).Value)
    If Len(specialReqs) > 0 Then
        If modRequirements.CheckRequirements(specialReqs) Then
            Dim specialDlg As String
            specialDlg = modUtils.SafeStr(ws.Cells(row, NPC_COL_DLG_SPECIAL).Value)
            If Len(specialDlg) > 0 Then
                GetDialogue = specialDlg
                Exit Function
            End If
        End If
    End If

    ' Select based on affinity
    Dim affinity As Long
    affinity = GetAffinity(npcID)

    If affinity >= AFFINITY_FRIENDLY Then
        Dim friendlyDlg As String
        friendlyDlg = modUtils.SafeStr(ws.Cells(row, NPC_COL_DLG_FRIENDLY).Value)
        If Len(friendlyDlg) > 0 Then
            GetDialogue = friendlyDlg
            Exit Function
        End If
    ElseIf affinity <= AFFINITY_HOSTILE Then
        Dim hostileDlg As String
        hostileDlg = modUtils.SafeStr(ws.Cells(row, NPC_COL_DLG_HOSTILE).Value)
        If Len(hostileDlg) > 0 Then
            GetDialogue = hostileDlg
            Exit Function
        End If
    End If

    ' Default dialogue
    GetDialogue = modUtils.SafeStr(ws.Cells(row, NPC_COL_DLG_DEFAULT).Value)
End Function

'===============================================================
' PUBLIC — Gift giving
'===============================================================

' Give an item to an NPC. Applies the NPC's gift effects and
' adjusts affinity. Returns True if the gift was accepted.
Public Function GiveGift(npcID As String, itemID As String) As Boolean
    GiveGift = False

    ' Must have the item
    If Not modInventory.HasItem(itemID) Then
        modUtils.DebugLog "modNPC.GiveGift: player doesn't have " & itemID
        Exit Function
    End If

    ' NPC must be alive
    If Not IsNPCAlive(npcID) Then Exit Function

    ' Remove item from inventory
    modInventory.RemoveItem itemID, 1

    ' Base affinity boost based on item value
    Dim itemValue As Long
    itemValue = modInventory.GetItemValue(itemID)
    Dim affinityGain As Long
    If itemValue >= 50 Then
        affinityGain = 15
    ElseIf itemValue >= 20 Then
        affinityGain = 10
    Else
        affinityGain = 5
    End If
    AdjustAffinity npcID, affinityGain

    ' Apply NPC-specific gift effects (if any)
    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row > 0 Then
        Dim ws As Worksheet
        Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
        If Not ws Is Nothing Then
            Dim giftEffects As String
            giftEffects = modUtils.SafeStr(ws.Cells(row, NPC_COL_GIFT_EFFECTS).Value)
            If Len(giftEffects) > 0 Then
                modEffects.ProcessEffects giftEffects
            End If
        End If
    End If

    GiveGift = True
    modUtils.DebugLog "modNPC.GiveGift: gave " & itemID & " to " & npcID & " (affinity +" & affinityGain & ")"
End Function

'===============================================================
' PUBLIC — NPC death / alive state
'===============================================================

' Kill an NPC
Public Sub KillNPC(npcID As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
    If ws Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetNPCRow(npcID)
    If row = 0 Then Exit Sub

    ws.Cells(row, NPC_COL_ALIVE).Value = False
    modUtils.DebugLog "modNPC.KillNPC: " & npcID & " is now dead"
End Sub

'===============================================================
' PUBLIC — Faction queries
'===============================================================

' Get all NPCs belonging to a faction. Returns a Collection of NPC IDs.
Public Function GetNPCsByFaction(faction As String) As Collection
    Dim result As New Collection

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_NPCS)
    If ws Is Nothing Then
        Set GetNPCsByFaction = result
        Exit Function
    End If

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, NPC_COL_ID)
        Dim npcID As String
        npcID = modUtils.SafeStr(ws.Cells(r, NPC_COL_ID).Value)
        If Len(npcID) = 0 Then GoTo NextFac

        If UCase(modUtils.SafeStr(ws.Cells(r, NPC_COL_FACTION).Value)) = UCase(faction) Then
            result.Add npcID
        End If
NextFac:
    Next r

    Set GetNPCsByFaction = result
End Function

' Get average suspicion across a faction (useful for hunter tracking)
Public Function GetFactionSuspicion(faction As String) As Long
    Dim npcs As Collection
    Set npcs = GetNPCsByFaction(faction)

    If npcs.Count = 0 Then
        GetFactionSuspicion = 0
        Exit Function
    End If

    Dim total As Long
    Dim i As Long
    For i = 1 To npcs.Count
        total = total + GetSuspicion(CStr(npcs(i)))
    Next i

    GetFactionSuspicion = total \ npcs.Count
End Function

' Adjust suspicion for all NPCs in a faction
Public Sub AdjustFactionSuspicion(faction As String, delta As Long)
    Dim npcs As Collection
    Set npcs = GetNPCsByFaction(faction)

    Dim i As Long
    For i = 1 To npcs.Count
        AdjustSuspicion CStr(npcs(i)), delta
    Next i

    modUtils.DebugLog "modNPC.AdjustFactionSuspicion: " & faction & " +" & delta
End Sub
