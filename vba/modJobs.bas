Attribute VB_Name = "modJobs"
'===============================================================
' modJobs — Job / Work System
' Damned Moon VBA RPG Engine — Phase 3
'===============================================================
' Lets the player take on jobs at locations that offer the JOB
' service. Jobs cost time, may have stat requirements, and pay
' out money, XP, items, and stat changes.
'
' Data source: tbl_Jobs
'   JobID, Name, Description, LocationFilter, TimeFilter,
'   TimeCost, Requirements, MoneyReward, XPReward, Effects,
'   Cooldown, RepeatFlag
'===============================================================

Option Explicit

' ── JOB TABLE COLUMN INDICES ──
Private Const JOB_COL_ID As Long = 1           ' A: JobID
Private Const JOB_COL_NAME As Long = 2         ' B: Display name
Private Const JOB_COL_DESC As Long = 3         ' C: Description / flavor text
Private Const JOB_COL_LOCATION As Long = 4     ' D: Location filter (pipe-delimited NodeIDs, or * for any)
Private Const JOB_COL_TIME As Long = 5         ' E: Time filter (pipe-delimited slots, or * for any)
Private Const JOB_COL_TIMECOST As Long = 6     ' F: Time cost in minutes
Private Const JOB_COL_REQS As Long = 7         ' G: Requirements to take the job
Private Const JOB_COL_MONEY As Long = 8        ' H: Money reward
Private Const JOB_COL_XP As Long = 9           ' I: XP reward
Private Const JOB_COL_EFFECTS As Long = 10     ' J: Effect string (same syntax as scene effects)
Private Const JOB_COL_COOLDOWN As Long = 11    ' K: Cooldown in days before repeatable (0 = no cooldown)
Private Const JOB_COL_REPEAT_FLAG As Long = 12 ' L: Flag set when completed (used for cooldown tracking)

'===============================================================
' PUBLIC — Get available jobs at a location
'===============================================================

' Returns a Collection of JobIDs available at the given node
' and current time of day. Filters by location, time, requirements,
' and cooldown.
Public Function GetAvailableJobs(nodeID As String) As Collection
    Dim result As New Collection

    Dim wsJobs As Worksheet
    Set wsJobs = modConfig.GetSheet(modConfig.SH_JOBS)
    If wsJobs Is Nothing Then
        Set GetAvailableJobs = result
        Exit Function
    End If

    Dim currentTime As String
    currentTime = UCase(modState.GetTimeOfDay())

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsJobs, JOB_COL_ID)
        Dim jobID As String
        jobID = modUtils.SafeStr(wsJobs.Cells(r, JOB_COL_ID).Value)
        If Len(jobID) = 0 Then GoTo NextJob

        ' Filter by location
        Dim locFilter As String
        locFilter = modUtils.SafeStr(wsJobs.Cells(r, JOB_COL_LOCATION).Value)
        If Len(locFilter) > 0 And locFilter <> "*" Then
            If Not MatchesFilter(nodeID, locFilter) Then GoTo NextJob
        End If

        ' Filter by time of day
        Dim timeFilter As String
        timeFilter = modUtils.SafeStr(wsJobs.Cells(r, JOB_COL_TIME).Value)
        If Len(timeFilter) > 0 And timeFilter <> "*" Then
            If Not MatchesFilter(currentTime, UCase(timeFilter)) Then GoTo NextJob
        End If

        ' Check requirements
        Dim reqs As String
        reqs = modUtils.SafeStr(wsJobs.Cells(r, JOB_COL_REQS).Value)
        If Len(reqs) > 0 Then
            If Not modRequirements.CheckRequirements(reqs) Then GoTo NextJob
        End If

        ' Check cooldown
        If IsJobOnCooldown(jobID, r) Then GoTo NextJob

        result.Add jobID
NextJob:
    Next r

    Set GetAvailableJobs = result
End Function

'===============================================================
' PUBLIC — Get job info
'===============================================================

' Get display name for a job
Public Function GetJobName(jobID As String) As String
    Dim row As Long
    row = modData.GetJobRow(jobID)
    If row = 0 Then
        GetJobName = jobID
        Exit Function
    End If
    GetJobName = modData.ReadCellStr(modConfig.SH_JOBS, row, JOB_COL_NAME)
    If Len(GetJobName) = 0 Then GetJobName = jobID
End Function

' Get description for a job
Public Function GetJobDescription(jobID As String) As String
    Dim row As Long
    row = modData.GetJobRow(jobID)
    If row = 0 Then Exit Function
    GetJobDescription = modData.ReadCellStr(modConfig.SH_JOBS, row, JOB_COL_DESC)
End Function

' Get time cost for a job in minutes
Public Function GetJobTimeCost(jobID As String) As Long
    Dim row As Long
    row = modData.GetJobRow(jobID)
    If row = 0 Then
        GetJobTimeCost = modTime.TIME_COST_JOB
        Exit Function
    End If
    GetJobTimeCost = modData.ReadCellLng(modConfig.SH_JOBS, row, JOB_COL_TIMECOST)
    If GetJobTimeCost <= 0 Then GetJobTimeCost = modTime.TIME_COST_JOB
End Function

' Get money reward for a job
Public Function GetJobMoney(jobID As String) As Long
    Dim row As Long
    row = modData.GetJobRow(jobID)
    If row = 0 Then Exit Function
    GetJobMoney = modData.ReadCellLng(modConfig.SH_JOBS, row, JOB_COL_MONEY)
End Function

' Get XP reward for a job
Public Function GetJobXP(jobID As String) As Long
    Dim row As Long
    row = modData.GetJobRow(jobID)
    If row = 0 Then Exit Function
    GetJobXP = modData.ReadCellLng(modConfig.SH_JOBS, row, JOB_COL_XP)
End Function

'===============================================================
' PUBLIC — Execute a job
'===============================================================

' Perform a job: spend time, check for success, apply rewards.
' Returns True if job completed successfully, False if blocked.
Public Function DoJob(jobID As String) As Boolean
    Dim row As Long
    row = modData.GetJobRow(jobID)
    If row = 0 Then
        modUtils.DebugLog "modJobs.DoJob: job '" & jobID & "' not found"
        DoJob = False
        Exit Function
    End If

    Dim wsJobs As Worksheet
    Set wsJobs = modConfig.GetSheet(modConfig.SH_JOBS)
    If wsJobs Is Nothing Then
        DoJob = False
        Exit Function
    End If

    ' Check requirements one more time
    Dim reqs As String
    reqs = modUtils.SafeStr(wsJobs.Cells(row, JOB_COL_REQS).Value)
    If Len(reqs) > 0 Then
        If Not modRequirements.CheckRequirements(reqs) Then
            modUtils.DebugLog "modJobs.DoJob: requirements not met for " & jobID
            DoJob = False
            Exit Function
        End If
    End If

    Application.ScreenUpdating = False

    ' Spend time
    Dim timeCost As Long
    timeCost = GetJobTimeCost(jobID)
    modTime.SpendTime timeCost

    ' Award money
    Dim moneyReward As Long
    moneyReward = modUtils.SafeLng(wsJobs.Cells(row, JOB_COL_MONEY).Value, 0)
    If moneyReward > 0 Then
        modState.AddStat modConfig.STAT_MONEY, moneyReward
    End If

    ' Award XP
    Dim xpReward As Long
    xpReward = modUtils.SafeLng(wsJobs.Cells(row, JOB_COL_XP).Value, 0)
    If xpReward > 0 Then
        modState.AddStat modConfig.STAT_XP, xpReward
    End If

    ' Apply job effects
    Dim effectStr As String
    effectStr = modUtils.SafeStr(wsJobs.Cells(row, JOB_COL_EFFECTS).Value)
    If Len(effectStr) > 0 Then
        modEffects.ProcessEffects effectStr
    End If

    ' Set completion flag for cooldown tracking
    Dim repeatFlag As String
    repeatFlag = modUtils.SafeStr(wsJobs.Cells(row, JOB_COL_REPEAT_FLAG).Value)
    If Len(repeatFlag) > 0 Then
        modState.SetFlag repeatFlag, True
    End If

    ' Build result narrative
    Dim jobName As String
    jobName = GetJobName(jobID)
    Dim resultText As String
    resultText = ChrW(&H2692) & " JOB COMPLETE: " & jobName & vbLf & vbLf
    resultText = resultText & "Time spent: " & timeCost & " minutes" & vbLf

    If moneyReward > 0 Then
        resultText = resultText & "Earned: $" & moneyReward & vbLf
    End If
    If xpReward > 0 Then
        resultText = resultText & "XP gained: " & xpReward & vbLf
    End If

    modUI.ShowNarrative resultText

    ' Update HUD
    modUI.UpdateStatsPanel
    modUI.UpdateDayTimePanel

    Application.ScreenUpdating = True

    modUtils.DebugLog "modJobs.DoJob: completed " & jobID & " ($" & moneyReward & ", " & xpReward & "XP, " & timeCost & "min)"
    DoJob = True
End Function

'===============================================================
' PUBLIC — Show job choices at current location
'===============================================================

' Populate the Game sheet with available job choices.
' Returns the number of jobs shown.
Public Function ShowJobChoices(nodeID As String) As Long
    Dim jobs As Collection
    Set jobs = GetAvailableJobs(nodeID)

    Dim count As Long
    count = 0

    Dim i As Long
    For i = 1 To jobs.Count
        If count >= modConfig.MAX_CHOICES Then Exit For

        Dim jobID As String
        jobID = CStr(jobs(i))

        Dim jobName As String
        jobName = GetJobName(jobID)

        Dim timeCost As Long
        timeCost = GetJobTimeCost(jobID)

        Dim money As Long
        money = GetJobMoney(jobID)

        Dim xp As Long
        xp = GetJobXP(jobID)

        ' Build choice label
        Dim label As String
        label = jobName & "  (" & timeCost & " min"
        If money > 0 Then label = label & ", $" & money
        If xp > 0 Then label = label & ", " & xp & "XP"
        label = label & ")"

        count = count + 1
        modUI.ShowChoiceButton count, CStr(count) & ".  " & label, True
    Next i

    ' Hide remaining buttons
    Dim j As Long
    For j = count + 1 To modConfig.MAX_CHOICES
        modUI.HideChoiceButton j
    Next j

    ShowJobChoices = count
End Function

' Build narrative text describing available work at a location
Public Function BuildJobNarrative(nodeID As String) As String
    Dim nodeName As String
    nodeName = modMap.GetNodeName(nodeID)

    Dim text As String
    text = "Work available at " & nodeName & ":" & vbLf & vbLf

    Dim jobs As Collection
    Set jobs = GetAvailableJobs(nodeID)

    If jobs.Count = 0 Then
        text = text & "No jobs available right now."
    Else
        Dim i As Long
        For i = 1 To jobs.Count
            Dim jobID As String
            jobID = CStr(jobs(i))
            Dim desc As String
            desc = GetJobDescription(jobID)
            If Len(desc) > 0 Then
                text = text & ChrW(&H25C6) & " " & GetJobName(jobID) & vbLf
                text = text & "   " & desc & vbLf & vbLf
            End If
        Next i
    End If

    text = text & vbLf & "What will you do?"
    BuildJobNarrative = text
End Function

'===============================================================
' PRIVATE — Cooldown check
'===============================================================

' Check if a job is on cooldown (completion flag set and not enough
' days have passed).
Private Function IsJobOnCooldown(jobID As String, row As Long) As Boolean
    IsJobOnCooldown = False

    Dim wsJobs As Worksheet
    Set wsJobs = modConfig.GetSheet(modConfig.SH_JOBS)
    If wsJobs Is Nothing Then Exit Function

    Dim repeatFlag As String
    repeatFlag = modUtils.SafeStr(wsJobs.Cells(row, JOB_COL_REPEAT_FLAG).Value)
    If Len(repeatFlag) = 0 Then Exit Function

    ' If the completion flag isn't set, not on cooldown
    If Not modState.GetFlag(repeatFlag) Then Exit Function

    ' Flag is set — check cooldown days
    Dim cooldownDays As Long
    cooldownDays = modUtils.SafeLng(wsJobs.Cells(row, JOB_COL_COOLDOWN).Value, 0)

    ' If cooldown is 0, it's a one-time job and flag blocks repeat
    If cooldownDays = 0 Then
        IsJobOnCooldown = True
        Exit Function
    End If

    ' For timed cooldowns, we'd need a "last completed day" tracker.
    ' For now, the flag system handles binary on/off cooldowns.
    ' A DAY_ADVANCE effect can clear repeat flags to re-enable jobs.
    IsJobOnCooldown = True
End Function

'===============================================================
' PRIVATE — Filter matching helper
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
