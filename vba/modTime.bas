Attribute VB_Name = "modTime"
'===============================================================
' modTime — Time & Day/Night Cycle System
' Damned Moon VBA RPG Engine — Phase 3
'===============================================================
' Manages time costs for actions, time-of-day gating, day/night
' gameplay effects, and periodic time-based events (hunger ticks,
' healing at rest, danger escalation at night).
'
' Builds on modState's raw time/day/moon storage by adding
' gameplay-level time logic.
'===============================================================

Option Explicit

' ── TIME SLOT ORDER (for comparison/progression) ──
Private Const SLOT_COUNT As Long = 6
' Index: 1=DAWN, 2=MORNING, 3=AFTERNOON, 4=DUSK, 5=NIGHT, 6=LATE_NIGHT

' ── DEFAULT TIME COSTS (minutes) ──
Public Const TIME_COST_TRAVEL_SHORT As Long = 30
Public Const TIME_COST_TRAVEL_MEDIUM As Long = 60
Public Const TIME_COST_TRAVEL_LONG As Long = 120
Public Const TIME_COST_REST As Long = 240          ' 4 hours
Public Const TIME_COST_FULL_REST As Long = 480     ' 8 hours (sleep)
Public Const TIME_COST_JOB As Long = 180           ' 3 hours
Public Const TIME_COST_SEARCH As Long = 60
Public Const TIME_COST_CONVERSATION As Long = 30

'===============================================================
' PUBLIC — Time slot utilities
'===============================================================

' Get the numeric index of a time slot (1-6). Returns 0 if unknown.
Public Function GetSlotIndex(timeSlot As String) As Long
    Select Case UCase(timeSlot)
        Case "DAWN": GetSlotIndex = 1
        Case "MORNING": GetSlotIndex = 2
        Case "AFTERNOON": GetSlotIndex = 3
        Case "DUSK": GetSlotIndex = 4
        Case "NIGHT": GetSlotIndex = 5
        Case "LATE_NIGHT": GetSlotIndex = 6
        Case Else: GetSlotIndex = 0
    End Select
End Function

' Get the time slot name from a numeric index (1-6)
Public Function GetSlotName(idx As Long) As String
    Select Case idx
        Case 1: GetSlotName = "DAWN"
        Case 2: GetSlotName = "MORNING"
        Case 3: GetSlotName = "AFTERNOON"
        Case 4: GetSlotName = "DUSK"
        Case 5: GetSlotName = "NIGHT"
        Case 6: GetSlotName = "LATE_NIGHT"
        Case Else: GetSlotName = "MORNING"
    End Select
End Function

' Get a display-friendly time slot name
Public Function GetSlotDisplayName(timeSlot As String) As String
    Select Case UCase(timeSlot)
        Case "DAWN": GetSlotDisplayName = "Dawn"
        Case "MORNING": GetSlotDisplayName = "Morning"
        Case "AFTERNOON": GetSlotDisplayName = "Afternoon"
        Case "DUSK": GetSlotDisplayName = "Dusk"
        Case "NIGHT": GetSlotDisplayName = "Night"
        Case "LATE_NIGHT": GetSlotDisplayName = "Late Night"
        Case Else: GetSlotDisplayName = timeSlot
    End Select
End Function

' Check if the given time slot is a "dark" period
Public Function IsDarkTime(timeSlot As String) As Boolean
    Dim s As String
    s = UCase(timeSlot)
    IsDarkTime = (s = "NIGHT" Or s = "LATE_NIGHT" Or s = "DUSK")
End Function

' Check if the current time is a dark period
Public Function IsCurrentlyDark() As Boolean
    IsCurrentlyDark = IsDarkTime(modState.GetTimeOfDay())
End Function

'===============================================================
' PUBLIC — Spend time on an action
'===============================================================

' Advance time by a number of minutes and apply any periodic effects
' that trigger during the elapsed time. Returns the new time slot.
Public Function SpendTime(minutes As Long) As String
    If minutes <= 0 Then
        SpendTime = modState.GetTimeOfDay()
        Exit Function
    End If

    Dim oldSlot As String
    oldSlot = modState.GetTimeOfDay()

    ' Advance the clock
    modState.AdvanceTime minutes

    Dim newSlot As String
    newSlot = modState.GetTimeOfDay()

    ' Apply periodic effects if the time slot changed
    If newSlot <> oldSlot Then
        ApplyTimeSlotChange oldSlot, newSlot
    End If

    SpendTime = newSlot
End Function

'===============================================================
' PUBLIC — Rest / Sleep
'===============================================================

' Rest for a given duration. Heals HP and reduces Rage.
' Returns the amount of HP healed.
Public Function Rest(minutes As Long) As Long
    Dim healAmount As Long

    ' Calculate healing based on rest duration
    If minutes >= TIME_COST_FULL_REST Then
        ' Full sleep: major healing
        healAmount = 30
    ElseIf minutes >= TIME_COST_REST Then
        ' Short rest: moderate healing
        healAmount = 15
    Else
        ' Quick break: minor healing
        healAmount = 5
    End If

    ' Apply healing
    modState.AddStat modConfig.STAT_HEALTH, healAmount

    ' Reduce rage from rest
    Dim rageReduction As Long
    rageReduction = minutes \ 60  ' 1 point per hour
    If rageReduction > 0 Then
        modState.AddStat modConfig.STAT_RAGE, -rageReduction
    End If

    ' Hunger increases while resting
    Dim hungerIncrease As Long
    hungerIncrease = minutes \ 120  ' 1 point per 2 hours
    If hungerIncrease > 0 Then
        modState.AddStat modConfig.STAT_HUNGER, hungerIncrease
    End If

    ' Advance time
    SpendTime minutes

    modUtils.DebugLog "modTime.Rest: " & minutes & "min, healed " & healAmount & "HP"
    Rest = healAmount
End Function

'===============================================================
' PUBLIC — Wait until a specific time slot
'===============================================================

' Wait until the target time slot. Advances time accordingly.
' Returns the number of minutes waited.
Public Function WaitUntil(targetSlot As String) As Long
    Dim currentIdx As Long
    currentIdx = GetSlotIndex(modState.GetTimeOfDay())
    Dim targetIdx As Long
    targetIdx = GetSlotIndex(targetSlot)

    If currentIdx = 0 Or targetIdx = 0 Then
        WaitUntil = 0
        Exit Function
    End If

    ' Calculate slots to advance
    Dim slotsAhead As Long
    If targetIdx > currentIdx Then
        slotsAhead = targetIdx - currentIdx
    ElseIf targetIdx < currentIdx Then
        slotsAhead = (SLOT_COUNT - currentIdx) + targetIdx
    Else
        ' Already at target slot
        WaitUntil = 0
        Exit Function
    End If

    ' Each slot is roughly 3-4 hours; approximate as minutes
    Dim minutesPerSlot As Long
    minutesPerSlot = 180  ' 3 hours average
    Dim totalMinutes As Long
    totalMinutes = slotsAhead * minutesPerSlot

    SpendTime totalMinutes

    modUtils.DebugLog "modTime.WaitUntil: waited " & totalMinutes & "min for " & targetSlot
    WaitUntil = totalMinutes
End Function

'===============================================================
' PUBLIC — Time-based danger multiplier
'===============================================================

' Returns a danger multiplier based on time of day and moon phase.
' 1.0 = normal, higher = more dangerous.
Public Function GetTimeDangerMultiplier() As Double
    Dim mult As Double
    mult = 1#

    ' Night is more dangerous
    If IsCurrentlyDark() Then
        mult = mult + 0.5
    End If

    ' Full moon dramatically increases danger
    If modState.IsFullMoon() Then
        mult = mult + 1#
    End If

    ' Late night is the most dangerous
    If UCase(modState.GetTimeOfDay()) = "LATE_NIGHT" Then
        mult = mult + 0.3
    End If

    GetTimeDangerMultiplier = mult
End Function

'===============================================================
' PUBLIC — Summary text for current time state
'===============================================================

' Build a display string for the current time, day, and moon
Public Function GetTimeSummary() As String
    Dim day As Long
    day = modState.GetCurrentDay()
    Dim tod As String
    tod = GetSlotDisplayName(modState.GetTimeOfDay())
    Dim moon As String
    moon = modState.GetMoonPhase()

    Dim summary As String
    summary = "Day " & day & ", " & tod
    If Len(moon) > 0 Then
        summary = summary & "  |  " & moon
    End If

    GetTimeSummary = summary
End Function

'===============================================================
' PRIVATE — Apply effects when time slot changes
'===============================================================

' Called whenever the time slot transitions (e.g., MORNING -> AFTERNOON).
' Applies periodic gameplay effects.
Private Sub ApplyTimeSlotChange(oldSlot As String, newSlot As String)
    ' Hunger ticks up each time slot
    modState.AddStat modConfig.STAT_HUNGER, 3

    ' Composure slowly recovers during daytime
    If Not IsDarkTime(newSlot) Then
        modState.AddStat modConfig.STAT_COMPOSURE, 2
    End If

    ' Rage creeps up at night
    If IsDarkTime(newSlot) Then
        modState.AddStat modConfig.STAT_RAGE, 5
    End If

    ' Instinct rises at night, falls during day
    If IsDarkTime(newSlot) Then
        modState.AddStat modConfig.STAT_INSTINCT, 3
    Else
        modState.AddStat modConfig.STAT_INSTINCT, -2
    End If

    ' Dawn recovery — slight HP heal if you survived the night
    If UCase(newSlot) = "DAWN" Then
        modState.AddStat modConfig.STAT_HEALTH, 5
        modState.AddStat modConfig.STAT_RAGE, -10
        modUtils.DebugLog "modTime: dawn recovery applied"
    End If

    modUtils.DebugLog "modTime.ApplyTimeSlotChange: " & oldSlot & " -> " & newSlot
End Sub
