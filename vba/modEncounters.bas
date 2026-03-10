Attribute VB_Name = "modEncounters"
'===============================================================
' modEncounters — Random Encounter System
' Damned Moon VBA RPG Engine — Phase 3
'===============================================================
' Rolls for random encounters during travel and exploration.
' Encounters are filtered by location, time of day, and moon
' phase, then selected by weighted random from eligible entries.
'
' Data source: tbl_Encounters
'   EncounterID, Name, Description, Type, LocationFilter,
'   TimeFilter, MoonFilter, Weight, DangerMin, Effects,
'   SceneJump, Requirements
'===============================================================

Option Explicit

' ── ENCOUNTER TABLE COLUMN INDICES ──
Private Const ENC_COL_ID As Long = 1           ' A: EncounterID
Private Const ENC_COL_NAME As Long = 2         ' B: Display name
Private Const ENC_COL_DESC As Long = 3         ' C: Narrative text
Private Const ENC_COL_TYPE As Long = 4         ' D: Type (TRAVEL, EXPLORE, REST, AMBIENT)
Private Const ENC_COL_LOCATION As Long = 5     ' E: Location filter (pipe-delimited NodeIDs, or * for any)
Private Const ENC_COL_TIME As Long = 6         ' F: Time filter (pipe-delimited slots, or * for any)
Private Const ENC_COL_MOON As Long = 7         ' G: Moon filter (keyword match, or * for any)
Private Const ENC_COL_WEIGHT As Long = 8       ' H: Base weight (higher = more likely)
Private Const ENC_COL_DANGER_MIN As Long = 9   ' I: Minimum danger level to trigger
Private Const ENC_COL_EFFECTS As Long = 10     ' J: Effect string (same syntax as scene effects)
Private Const ENC_COL_SCENEJUMP As Long = 11   ' K: Scene to jump to (optional)
Private Const ENC_COL_REQS As Long = 12        ' L: Requirements (optional)

'===============================================================
' PUBLIC — Roll for a random encounter during travel
'===============================================================

' Roll for an encounter while traveling from one node to another.
' effectiveDanger is the combined route + destination + time danger.
' Returns the EncounterID if triggered, or "" for no encounter.
Public Function RollEncounter(fromNodeID As String, toNodeID As String, effectiveDanger As Long) As String
    RollEncounter = ""

    ' Base encounter chance: danger / 200 (so danger 100 = 50% chance)
    Dim encounterChance As Long
    encounterChance = effectiveDanger

    ' Apply time-based danger multiplier
    Dim timeMult As Double
    timeMult = modTime.GetTimeDangerMultiplier()
    encounterChance = CLng(encounterChance * timeMult)

    ' Roll d100
    Dim roll As Long
    roll = modUtils.RandBetween(1, 100)

    ' No encounter if roll exceeds threshold
    If roll > modUtils.Clamp(encounterChance, 5, 85) Then
        modUtils.DebugLog "modEncounters.RollEncounter: no encounter (roll=" & roll & ", threshold=" & encounterChance & ")"
        Exit Function
    End If

    ' Build pool of eligible encounters
    Dim eligible As Collection
    Set eligible = GetEligibleEncounters("TRAVEL", fromNodeID, toNodeID, effectiveDanger)

    If eligible.Count = 0 Then
        modUtils.DebugLog "modEncounters.RollEncounter: no eligible encounters"
        Exit Function
    End If

    ' Weighted selection
    RollEncounter = SelectWeightedEncounter(eligible)
    modUtils.DebugLog "modEncounters.RollEncounter: triggered " & RollEncounter
End Function

' Roll for an encounter at a location (exploration, resting, etc.)
' encounterType: "EXPLORE", "REST", "AMBIENT"
Public Function RollLocationEncounter(nodeID As String, encounterType As String) As String
    RollLocationEncounter = ""

    Dim danger As Long
    danger = 0

    ' Get node danger if available
    If modData.MapNodeExists(nodeID) Then
        danger = modMap.GetNodeDanger(nodeID)
    End If

    ' Apply time multiplier
    Dim timeMult As Double
    timeMult = modTime.GetTimeDangerMultiplier()
    danger = CLng(danger * timeMult)

    ' Roll d100
    Dim roll As Long
    roll = modUtils.RandBetween(1, 100)

    If roll > modUtils.Clamp(danger, 5, 70) Then
        Exit Function
    End If

    ' Build eligible pool
    Dim eligible As Collection
    Set eligible = GetEligibleEncounters(encounterType, nodeID, "", danger)

    If eligible.Count = 0 Then Exit Function

    RollLocationEncounter = SelectWeightedEncounter(eligible)
    modUtils.DebugLog "modEncounters.RollLocationEncounter: triggered " & RollLocationEncounter & " at " & nodeID
End Function

'===============================================================
' PUBLIC — Run (resolve) an encounter
'===============================================================

' Execute an encounter: show its narrative, apply effects, and
' optionally jump to a scene.
Public Sub RunEncounter(encounterID As String)
    Dim wsEnc As Worksheet
    Set wsEnc = modConfig.GetSheet(modConfig.SH_ENCOUNTERS)
    If wsEnc Is Nothing Then Exit Sub

    Dim row As Long
    row = modData.GetEncounterRow(encounterID)
    If row = 0 Then
        modUtils.DebugLog "modEncounters.RunEncounter: encounter '" & encounterID & "' not found"
        Exit Sub
    End If

    ' Get encounter data
    Dim encName As String
    encName = modUtils.SafeStr(wsEnc.Cells(row, ENC_COL_NAME).Value)
    Dim encDesc As String
    encDesc = modUtils.SafeStr(wsEnc.Cells(row, ENC_COL_DESC).Value)
    Dim effectStr As String
    effectStr = modUtils.SafeStr(wsEnc.Cells(row, ENC_COL_EFFECTS).Value)
    Dim sceneJump As String
    sceneJump = modUtils.SafeStr(wsEnc.Cells(row, ENC_COL_SCENEJUMP).Value)

    ' Show encounter narrative as an interstitial
    If Len(encDesc) > 0 Then
        Dim fullText As String
        fullText = ChrW(&H26A0) & " " & encName & vbLf & vbLf & encDesc
        modUI.ShowNarrative fullText
    End If

    ' Apply encounter effects
    If Len(effectStr) > 0 Then
        Dim jumpFromEffects As String
        jumpFromEffects = modEffects.ProcessEffects(effectStr)
        ' Effect-triggered jump takes priority
        If Len(jumpFromEffects) > 0 And Len(sceneJump) = 0 Then
            sceneJump = jumpFromEffects
        End If
    End If

    ' Update UI after effects
    modUI.UpdateStatsPanel

    ' Jump to scene if specified
    If Len(sceneJump) > 0 Then
        modSceneEngine.LoadScene sceneJump
    End If

    modUtils.DebugLog "modEncounters.RunEncounter: resolved " & encounterID
End Sub

'===============================================================
' PUBLIC — Get encounter info
'===============================================================

' Get the display name for an encounter
Public Function GetEncounterName(encounterID As String) As String
    Dim row As Long
    row = modData.GetEncounterRow(encounterID)
    If row = 0 Then Exit Function
    GetEncounterName = modData.ReadCellStr(modConfig.SH_ENCOUNTERS, row, ENC_COL_NAME)
End Function

' Get the description for an encounter
Public Function GetEncounterDescription(encounterID As String) As String
    Dim row As Long
    row = modData.GetEncounterRow(encounterID)
    If row = 0 Then Exit Function
    GetEncounterDescription = modData.ReadCellStr(modConfig.SH_ENCOUNTERS, row, ENC_COL_DESC)
End Function

'===============================================================
' PRIVATE — Build eligible encounter pool
'===============================================================

' Returns a Collection of arrays: Array(EncounterID, Weight)
' filtered by type, location, time, moon, danger, and requirements.
Private Function GetEligibleEncounters(encounterType As String, _
                                        nodeID As String, _
                                        Optional toNodeID As String = "", _
                                        Optional effectiveDanger As Long = 0) As Collection
    Dim result As New Collection

    Dim wsEnc As Worksheet
    Set wsEnc = modConfig.GetSheet(modConfig.SH_ENCOUNTERS)
    If wsEnc Is Nothing Then
        Set GetEligibleEncounters = result
        Exit Function
    End If

    Dim currentTime As String
    currentTime = UCase(modState.GetTimeOfDay())
    Dim currentMoon As String
    currentMoon = UCase(modState.GetMoonPhase())

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsEnc, ENC_COL_ID)
        Dim encID As String
        encID = modUtils.SafeStr(wsEnc.Cells(r, ENC_COL_ID).Value)
        If Len(encID) = 0 Then GoTo NextEnc

        ' Filter by type
        Dim encType As String
        encType = UCase(modUtils.SafeStr(wsEnc.Cells(r, ENC_COL_TYPE).Value))
        If Len(encType) > 0 And encType <> "*" Then
            If encType <> UCase(encounterType) Then GoTo NextEnc
        End If

        ' Filter by location
        Dim locFilter As String
        locFilter = modUtils.SafeStr(wsEnc.Cells(r, ENC_COL_LOCATION).Value)
        If Len(locFilter) > 0 And locFilter <> "*" Then
            If Not MatchesFilter(nodeID, locFilter) Then
                ' Also check destination for travel encounters
                If Len(toNodeID) = 0 Or Not MatchesFilter(toNodeID, locFilter) Then
                    GoTo NextEnc
                End If
            End If
        End If

        ' Filter by time
        Dim timeFilter As String
        timeFilter = modUtils.SafeStr(wsEnc.Cells(r, ENC_COL_TIME).Value)
        If Len(timeFilter) > 0 And timeFilter <> "*" Then
            If Not MatchesFilter(currentTime, UCase(timeFilter)) Then GoTo NextEnc
        End If

        ' Filter by moon
        Dim moonFilter As String
        moonFilter = modUtils.SafeStr(wsEnc.Cells(r, ENC_COL_MOON).Value)
        If Len(moonFilter) > 0 And moonFilter <> "*" Then
            If InStr(currentMoon, UCase(moonFilter)) = 0 Then GoTo NextEnc
        End If

        ' Filter by minimum danger
        Dim dangerMin As Long
        dangerMin = modUtils.SafeLng(wsEnc.Cells(r, ENC_COL_DANGER_MIN).Value, 0)
        If effectiveDanger < dangerMin Then GoTo NextEnc

        ' Filter by requirements
        Dim reqs As String
        reqs = modUtils.SafeStr(wsEnc.Cells(r, ENC_COL_REQS).Value)
        If Len(reqs) > 0 Then
            If Not modRequirements.CheckRequirements(reqs) Then GoTo NextEnc
        End If

        ' Passed all filters — add to eligible pool
        Dim weight As Long
        weight = modUtils.SafeLng(wsEnc.Cells(r, ENC_COL_WEIGHT).Value, 10)
        If weight <= 0 Then weight = 10

        result.Add Array(encID, weight)
NextEnc:
    Next r

    Set GetEligibleEncounters = result
End Function

'===============================================================
' PRIVATE — Weighted random selection from eligible pool
'===============================================================
Private Function SelectWeightedEncounter(eligible As Collection) As String
    If eligible.Count = 0 Then
        SelectWeightedEncounter = ""
        Exit Function
    End If

    ' Build weight collection for WeightedPick
    Dim weights As New Collection
    Dim i As Long
    For i = 1 To eligible.Count
        Dim entry As Variant
        entry = eligible(i)
        weights.Add CLng(entry(1))
    Next i

    Dim picked As Long
    picked = modUtils.WeightedPick(weights)

    If picked > 0 And picked <= eligible.Count Then
        Dim pickedEntry As Variant
        pickedEntry = eligible(picked)
        SelectWeightedEncounter = CStr(pickedEntry(0))
    End If
End Function

'===============================================================
' PRIVATE — Check if a value matches a pipe-delimited filter
'===============================================================
Private Function MatchesFilter(val As String, filterStr As String) As Boolean
    If filterStr = "*" Then
        MatchesFilter = True
        Exit Function
    End If

    Dim parts As Variant
    parts = modUtils.SplitTrimmed(filterStr, "|")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If UCase(Trim(CStr(parts(i)))) = UCase(val) Then
            MatchesFilter = True
            Exit Function
        End If
    Next i

    MatchesFilter = False
End Function
