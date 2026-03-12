Attribute VB_Name = "modState"
'===============================================================
' modState — Live Game State Manager
' Damned Moon VBA RPG Engine
'===============================================================
' All state changes flow through this module. Stats, flags,
' current scene, location, time, moon, equipped items — all of it.
' No random writes to cells from gremlins elsewhere.
'===============================================================

Option Explicit

'===============================================================
' STAT SYSTEM — Get / Set / Add
'===============================================================

' Get a numeric stat value by name
Public Function GetStat(statName As String) As Long
    Dim row As Long
    row = modData.GetStatRow(statName)
    If row = 0 Then
        GetStat = 0
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_STATS)
    If ws Is Nothing Then
        GetStat = 0
        Exit Function
    End If

    GetStat = modUtils.SafeLng(ws.Cells(row, 3).Value, 0)
End Function

' Get a text-valued stat (e.g., TIME_OF_DAY, MOON_PHASE)
Public Function GetStatText(statName As String) As String
    Dim row As Long
    row = modData.GetStatRow(statName)
    If row = 0 Then
        GetStatText = ""
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_STATS)
    If ws Is Nothing Then
        GetStatText = ""
        Exit Function
    End If

    GetStatText = modUtils.SafeStr(ws.Cells(row, 3).Value)
End Function

' Set a stat to a specific value (numeric or text)
Public Sub SetStat(statName As String, val As Variant)
    Dim row As Long
    row = modData.GetStatRow(statName)
    If row = 0 Then
        modUtils.DebugLog "modState.SetStat: stat '" & statName & "' not found in cache"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_STATS)
    If ws Is Nothing Then Exit Sub

    ws.Cells(row, 3).Value = val
    modUtils.DebugLog "modState.SetStat: " & statName & " = " & CStr(val)
End Sub

' Add a delta to a numeric stat (clamped for core stats)
Public Sub AddStat(ByVal statName As String, ByVal delta As Double)
    Dim current As Long
    current = GetStat(statName)
    current = current + CLng(delta)

    ' Clamp core stats to 0-100
    If IsCoreStatName(statName) Then
        current = modUtils.Clamp(current, 0, 100)
    End If

    SetStat statName, current
End Sub

' Clamp all core stats to valid ranges
Public Sub ClampStats()
    Dim coreNames As Variant
    coreNames = Split(modConfig.CORE_STATS, ",")

    Dim i As Long
    For i = LBound(coreNames) To UBound(coreNames)
        Dim sn As String
        sn = Trim(CStr(coreNames(i)))
        Dim v As Long
        v = GetStat(sn)
        v = modUtils.Clamp(v, 0, 100)
        SetStat sn, v
    Next i
End Sub

' Check if a stat name is a core stat (0-100 clamped)
Private Function IsCoreStatName(statName As String) As Boolean
    IsCoreStatName = (InStr("," & modConfig.CORE_STATS & ",", "," & statName & ",") > 0)
End Function

'===============================================================
' FLAG SYSTEM — Get / Set / Toggle
'===============================================================

' Get a boolean flag value
Public Function GetFlag(flagName As String) As Boolean
    Dim row As Long
    row = modData.GetFlagRow(flagName)
    If row = 0 Then
        GetFlag = False
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_FLAGS)
    If ws Is Nothing Then
        GetFlag = False
        Exit Function
    End If

    GetFlag = modUtils.SafeBool(ws.Cells(row, 2).Value, False)
End Function

' Set a flag to True or False
Public Sub SetFlag(flagName As String, val As Boolean)
    Dim row As Long
    row = modData.GetFlagRow(flagName)
    If row = 0 Then
        modUtils.DebugLog "modState.SetFlag: flag '" & flagName & "' not found in cache"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_FLAGS)
    If ws Is Nothing Then Exit Sub

    ws.Cells(row, 2).Value = val
    modUtils.DebugLog "modState.SetFlag: " & flagName & " = " & CStr(val)
End Sub

' Toggle a flag
Public Sub ToggleFlag(flagName As String)
    SetFlag flagName, Not GetFlag(flagName)
End Sub

'===============================================================
' CURRENT SCENE
'===============================================================

' Get the current scene ID from the Game sheet
Public Function GetCurrentScene() As String
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then
        GetCurrentScene = ""
        Exit Function
    End If
    GetCurrentScene = modUtils.SafeStr(ws.Range(modConfig.SCENE_ID_CELL).Value)
End Function

' Set the current scene ID on the Game sheet
Public Sub SetCurrentScene(sceneID As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub
    ws.Range(modConfig.SCENE_ID_CELL).Value = sceneID
    modUtils.DebugLog "modState.SetCurrentScene: " & sceneID
End Sub

'===============================================================
' CURRENT LOCATION
'===============================================================

' Get the current location code
Public Function GetCurrentLocation() As String
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then
        GetCurrentLocation = ""
        Exit Function
    End If
    GetCurrentLocation = modUtils.SafeStr(ws.Range(modConfig.LOCATION_CELL).Value)
End Function

' Set the current location code
Public Sub SetCurrentLocation(nodeID As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub
    ws.Range(modConfig.LOCATION_CELL).Value = nodeID
    modUtils.DebugLog "modState.SetCurrentLocation: " & nodeID
End Sub

'===============================================================
' DAY / TIME / MOON
'===============================================================

' Get current day number
Public Function GetCurrentDay() As Long
    GetCurrentDay = GetStat(modConfig.STAT_DAY_COUNTER)
End Function

' Set current day number
Public Sub SetCurrentDay(dayNum As Long)
    SetStat modConfig.STAT_DAY_COUNTER, dayNum
End Sub

' Get current time of day (text: DAWN, MORNING, AFTERNOON, DUSK, NIGHT, LATE_NIGHT)
Public Function GetTimeOfDay() As String
    GetTimeOfDay = GetStatText(modConfig.STAT_TIME_OF_DAY)
End Function

' Set time of day
Public Sub SetTimeOfDay(timeSlot As String)
    SetStat modConfig.STAT_TIME_OF_DAY, timeSlot
End Sub

' Advance time by a number of minutes (updates time slot based on thresholds)
Public Sub AdvanceTime(minutes As Long)
    ' Time slots: DAWN(5-7), MORNING(7-12), AFTERNOON(12-17), DUSK(17-19), NIGHT(19-24), LATE_NIGHT(0-5)
    ' For now, store total minutes and derive time slot
    Dim totalMin As Long
    totalMin = GetStat("TIME_MINUTES") + minutes

    ' Each day is 1440 minutes
    If totalMin >= 1440 Then
        totalMin = totalMin - 1440
        AdvanceDay
    End If

    SetStat "TIME_MINUTES", totalMin

    ' Derive time of day from minutes
    Dim hours As Long
    hours = totalMin \ 60

    Dim newSlot As String
    Select Case hours
        Case 5, 6: newSlot = "DAWN"
        Case 7 To 11: newSlot = "MORNING"
        Case 12 To 16: newSlot = "AFTERNOON"
        Case 17, 18: newSlot = "DUSK"
        Case 19 To 23: newSlot = "NIGHT"
        Case Else: newSlot = "LATE_NIGHT"
    End Select

    SetTimeOfDay newSlot
    modUtils.DebugLog "modState.AdvanceTime: +" & minutes & "min -> " & newSlot
End Sub

' Advance to the next day
Public Sub AdvanceDay()
    Dim currentDay As Long
    currentDay = GetCurrentDay()
    SetCurrentDay currentDay + 1
    TickMoonPhase
    modUtils.DebugLog "modState.AdvanceDay: day " & (currentDay + 1)
End Sub

' Get the current moon phase name
Public Function GetMoonPhase() As String
    GetMoonPhase = GetStatText(modConfig.STAT_MOON_PHASE)
End Function

' Tick the moon phase forward (called on day advance)
Public Sub TickMoonPhase()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_MOON)
    If ws Is Nothing Then Exit Sub

    Dim day As Long
    day = GetCurrentDay()

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, 1)
        Dim dayRange As String
        dayRange = modUtils.SafeStr(ws.Cells(r, 2).Value)

        If InStr(dayRange, "-") > 0 Then
            Dim parts() As String
            parts = Split(dayRange, "-")
            If day >= CLng(parts(0)) And day <= CLng(parts(1)) Then
                SetStat modConfig.STAT_MOON_PHASE, CStr(ws.Cells(r, 1).Value)
                Exit Sub
            End If
        ElseIf IsNumeric(dayRange) Then
            If day = CLng(dayRange) Then
                SetStat modConfig.STAT_MOON_PHASE, CStr(ws.Cells(r, 1).Value)
                Exit Sub
            End If
        End If
    Next r
End Sub

' Check if it is currently nighttime
Public Function IsNight() As Boolean
    Dim tod As String
    tod = GetTimeOfDay()
    IsNight = (tod = "NIGHT" Or tod = "LATE_NIGHT" Or tod = "DUSK")
End Function

' Check if we're in a full moon
Public Function IsFullMoon() As Boolean
    Dim phase As String
    phase = GetMoonPhase()
    IsFullMoon = (InStr(UCase(phase), "FULL") > 0)
End Function

'===============================================================
' EQUIPPED ITEM TRACKING
'===============================================================

' Get the currently equipped weapon item ID (or "" if none)
Public Function GetEquippedWeapon() As String
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_INV)
    If ws Is Nothing Then
        GetEquippedWeapon = ""
        Exit Function
    End If

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, 1)
        If modUtils.SafeBool(ws.Cells(r, 5).Value) Then
            GetEquippedWeapon = modUtils.SafeStr(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
    GetEquippedWeapon = ""
End Function

'===============================================================
' DERIVED STATS
'===============================================================

' Control = Humanity - Rage (clamped 0-100)
Public Function GetControl() As Long
    GetControl = modUtils.Clamp(GetStat(modConfig.STAT_HUMANITY) - GetStat(modConfig.STAT_RAGE), 0, 100)
End Function

' Danger level — higher means more risk of transformation / combat penalty
Public Function GetDangerLevel() As Long
    Dim rage As Long: rage = GetStat(modConfig.STAT_RAGE)
    Dim hunger As Long: hunger = GetStat(modConfig.STAT_HUNGER)
    Dim humanity As Long: humanity = GetStat(modConfig.STAT_HUMANITY)
    GetDangerLevel = modUtils.Clamp((rage + hunger) \ 2 - humanity \ 4, 0, 100)
End Function

'===============================================================
' FULL STATE RESET — used by NewGame
'===============================================================
Public Sub ResetAllStats()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_STATS)
    If ws Is Nothing Then Exit Sub

    ' Column B = base/default value, Column C = current value
    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, 1)
        ws.Cells(r, 3).Value = ws.Cells(r, 2).Value
    Next r

    modUtils.DebugLog "modState.ResetAllStats: all stats reset to base values"
End Sub

Public Sub ResetAllFlags()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_FLAGS)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, 1)
        ws.Cells(r, 2).Value = False
    Next r

    modUtils.DebugLog "modState.ResetAllFlags: all flags cleared"
End Sub

Public Sub ResetInventory()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_INV)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, 1)
        ws.Cells(r, 2).Value = ""   ' ItemID
        ws.Cells(r, 3).Value = ""   ' ItemName
        ws.Cells(r, 4).Value = 0    ' Qty
        ws.Cells(r, 5).Value = False ' Equipped
    Next r

    modUtils.DebugLog "modState.ResetInventory: inventory cleared"
End Sub

Public Sub ResetQuests()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, 1)
        Dim qType As String
        qType = modUtils.SafeStr(ws.Cells(r, 3).Value)
        If qType = "MAIN" Then
            ws.Cells(r, 4).Value = "ACTIVE"
            ws.Cells(r, 6).Value = 0
        Else
            ws.Cells(r, 4).Value = "INACTIVE"
            ws.Cells(r, 6).Value = -1
        End If
    Next r

    modUtils.DebugLog "modState.ResetQuests: quests reset"
End Sub

' Reset everything: stats, flags, inventory, quests
Public Sub ResetGameState()
    ResetAllStats
    ResetAllFlags
    ResetInventory
    ResetQuests
    SetCurrentScene ""
    SetCurrentLocation modConfig.DEFAULT_START_LOCATION
    modUtils.DebugLog "modState.ResetGameState: full reset complete"
End Sub
