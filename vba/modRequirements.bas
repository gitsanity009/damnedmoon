Attribute VB_Name = "modRequirements"
'===============================================================
' modRequirements — Requirement Checking for Choices
' Damned Moon VBA RPG Engine — Phase 2
'===============================================================
' Evaluates requirement strings from the scene table to determine
' whether a choice is available to the player. Supports flags,
' stat comparisons, item checks, and compound requirements.
'
' Requirement syntax (pipe-delimited for multiple):
'   FLAG:FlagName            — flag must be TRUE
'   !FLAG:FlagName           — flag must be FALSE
'   STAT:RAGE>50             — stat comparison (>, <, =, >=, <=)
'   ITEM:ITM_SILVER_KNIFE    — player has item in inventory
'   !ITEM:ITM_SILVER_KNIFE   — player does NOT have item
'   TIME:NIGHT               — current time of day matches
'   MOON:FULL                — moon phase contains keyword
'   LOCATION:NODE_ID         — player is at a specific map node
'   MONEY:50                 — player has >= N money
'===============================================================

Option Explicit

'===============================================================
' PUBLIC — Check a full requirement string (may be compound)
'===============================================================

' Check if ALL requirements in a pipe-delimited string are met.
' Returns True if reqStr is empty (no requirements).
Public Function CheckRequirements(reqStr As String) As Boolean
    If Len(Trim(reqStr)) = 0 Then
        CheckRequirements = True
        Exit Function
    End If

    Dim reqs As Variant
    reqs = modUtils.SplitTrimmed(reqStr, modConfig.EFFECT_DELIM)

    Dim i As Long
    For i = LBound(reqs) To UBound(reqs)
        Dim req As String
        req = Trim(CStr(reqs(i)))
        If Len(req) = 0 Then GoTo NextReq

        If Not CheckSingleRequirement(req) Then
            CheckRequirements = False
            Exit Function
        End If
NextReq:
    Next i

    CheckRequirements = True
End Function

' Get a human-readable reason why a requirement is not met.
' Returns "" if all requirements are met.
Public Function GetFailReason(reqStr As String) As String
    If Len(Trim(reqStr)) = 0 Then
        GetFailReason = ""
        Exit Function
    End If

    Dim reqs As Variant
    reqs = modUtils.SplitTrimmed(reqStr, modConfig.EFFECT_DELIM)

    Dim i As Long
    For i = LBound(reqs) To UBound(reqs)
        Dim req As String
        req = Trim(CStr(reqs(i)))
        If Len(req) = 0 Then GoTo NextReq2

        If Not CheckSingleRequirement(req) Then
            GetFailReason = DescribeRequirement(req)
            Exit Function
        End If
NextReq2:
    Next i

    GetFailReason = ""
End Function

'===============================================================
' PRIVATE — Single requirement evaluation
'===============================================================

Private Function CheckSingleRequirement(req As String) As Boolean
    Dim negated As Boolean
    Dim token As String
    token = req

    ' Check for negation prefix
    If Left(token, 1) = "!" Then
        negated = True
        token = Mid(token, 2)
    End If

    Dim result As Boolean

    If modUtils.StartsWith(token, "FLAG:") Then
        result = EvalFlagReq(modUtils.StripPrefix(token, "FLAG:"))

    ElseIf modUtils.StartsWith(token, "STAT:") Then
        result = EvalStatReq(modUtils.StripPrefix(token, "STAT:"))

    ElseIf modUtils.StartsWith(token, "ITEM:") Then
        result = EvalItemReq(modUtils.StripPrefix(token, "ITEM:"))

    ElseIf modUtils.StartsWith(token, "TIME:") Then
        result = EvalTimeReq(modUtils.StripPrefix(token, "TIME:"))

    ElseIf modUtils.StartsWith(token, "MOON:") Then
        result = EvalMoonReq(modUtils.StripPrefix(token, "MOON:"))

    ElseIf modUtils.StartsWith(token, "LOCATION:") Then
        result = EvalLocationReq(modUtils.StripPrefix(token, "LOCATION:"))

    ElseIf modUtils.StartsWith(token, "MONEY:") Then
        result = EvalMoneyReq(modUtils.StripPrefix(token, "MONEY:"))

    Else
        ' Unknown requirement type — treat as met (permissive)
        modUtils.DebugLog "modRequirements: unknown req type '" & req & "', treating as met"
        result = True
    End If

    If negated Then
        CheckSingleRequirement = Not result
    Else
        CheckSingleRequirement = result
    End If
End Function

'===============================================================
' FLAG REQUIREMENT: FLAG:FlagName
'===============================================================
Private Function EvalFlagReq(flagName As String) As Boolean
    EvalFlagReq = modState.GetFlag(flagName)
End Function

'===============================================================
' STAT REQUIREMENT: STAT:RAGE>50, STAT:HUMANITY>=30
'===============================================================
Private Function EvalStatReq(expr As String) As Boolean
    ' Find comparison operator position
    Dim opPos As Long
    Dim opLen As Long
    Dim op As String

    opPos = 0
    opLen = 1

    Dim k As Long
    For k = 1 To Len(expr)
        Dim ch As String
        ch = Mid(expr, k, 1)
        If ch = ">" Or ch = "<" Or ch = "=" Then
            opPos = k
            ' Check for two-char operators (>=, <=)
            If k < Len(expr) Then
                Dim nextCh As String
                nextCh = Mid(expr, k + 1, 1)
                If nextCh = "=" Then
                    opLen = 2
                End If
            End If
            Exit For
        End If
    Next k

    If opPos = 0 Then
        ' No operator — just check if stat is > 0
        EvalStatReq = (modState.GetStat(expr) > 0)
        Exit Function
    End If

    Dim statName As String
    statName = Left(expr, opPos - 1)
    op = Mid(expr, opPos, opLen)
    Dim targetVal As Long
    targetVal = modUtils.SafeLng(Mid(expr, opPos + opLen), 0)

    Dim current As Long
    current = modState.GetStat(statName)

    Select Case op
        Case ">": EvalStatReq = (current > targetVal)
        Case "<": EvalStatReq = (current < targetVal)
        Case "=": EvalStatReq = (current = targetVal)
        Case ">=": EvalStatReq = (current >= targetVal)
        Case "<=": EvalStatReq = (current <= targetVal)
        Case Else: EvalStatReq = True
    End Select
End Function

'===============================================================
' ITEM REQUIREMENT: ITEM:ITM_SILVER_KNIFE
'===============================================================
Private Function EvalItemReq(itemID As String) As Boolean
    ' Check if item exists in inventory with qty > 0
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_INV)
    If ws Is Nothing Then
        EvalItemReq = False
        Exit Function
    End If

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(ws, 1)
        If modUtils.SafeStr(ws.Cells(r, 2).Value) = itemID Then
            Dim qty As Long
            qty = modUtils.SafeLng(ws.Cells(r, 4).Value, 0)
            EvalItemReq = (qty > 0)
            Exit Function
        End If
    Next r

    EvalItemReq = False
End Function

'===============================================================
' TIME REQUIREMENT: TIME:NIGHT, TIME:DAWN
'===============================================================
Private Function EvalTimeReq(timeVal As String) As Boolean
    Dim current As String
    current = UCase(modState.GetTimeOfDay())
    EvalTimeReq = (current = UCase(timeVal))
End Function

'===============================================================
' MOON REQUIREMENT: MOON:FULL, MOON:NEW
'===============================================================
Private Function EvalMoonReq(moonVal As String) As Boolean
    Dim current As String
    current = UCase(modState.GetMoonPhase())
    EvalMoonReq = (InStr(current, UCase(moonVal)) > 0)
End Function

'===============================================================
' LOCATION REQUIREMENT: LOCATION:NODE_ID
'===============================================================
Private Function EvalLocationReq(locID As String) As Boolean
    EvalLocationReq = (UCase(modState.GetCurrentLocation()) = UCase(locID))
End Function

'===============================================================
' MONEY REQUIREMENT: MONEY:50
'===============================================================
Private Function EvalMoneyReq(expr As String) As Boolean
    ' Supports comparison operators: MONEY:>=50, MONEY:50 (treated as >=)
    Dim opPos As Long
    opPos = modUtils.FindComparisonPos(expr)

    If opPos > 0 Then
        ' Has operator — delegate to stat-style eval
        EvalMoneyReq = EvalStatReq("MONEY" & expr)
    Else
        ' Bare number = must have >= that amount
        Dim needed As Long
        needed = modUtils.SafeLng(expr, 0)
        EvalMoneyReq = (modState.GetStat(modConfig.STAT_MONEY) >= needed)
    End If
End Function

'===============================================================
' DESCRIBE — Human-readable requirement description
'===============================================================
Private Function DescribeRequirement(req As String) As String
    Dim negated As Boolean
    Dim token As String
    token = req

    If Left(token, 1) = "!" Then
        negated = True
        token = Mid(token, 2)
    End If

    If modUtils.StartsWith(token, "FLAG:") Then
        Dim fn As String
        fn = modUtils.StripPrefix(token, "FLAG:")
        If negated Then
            DescribeRequirement = "[Requires: " & fn & " not set]"
        Else
            DescribeRequirement = "[Requires: " & fn & "]"
        End If

    ElseIf modUtils.StartsWith(token, "STAT:") Then
        DescribeRequirement = "[Requires: " & modUtils.StripPrefix(token, "STAT:") & "]"

    ElseIf modUtils.StartsWith(token, "ITEM:") Then
        Dim iid As String
        iid = modUtils.StripPrefix(token, "ITEM:")
        If negated Then
            DescribeRequirement = "[Must not have: " & iid & "]"
        Else
            DescribeRequirement = "[Requires item: " & iid & "]"
        End If

    ElseIf modUtils.StartsWith(token, "TIME:") Then
        DescribeRequirement = "[Requires time: " & modUtils.StripPrefix(token, "TIME:") & "]"

    ElseIf modUtils.StartsWith(token, "MOON:") Then
        DescribeRequirement = "[Requires moon: " & modUtils.StripPrefix(token, "MOON:") & "]"

    ElseIf modUtils.StartsWith(token, "LOCATION:") Then
        Dim lid As String
        lid = modUtils.StripPrefix(token, "LOCATION:")
        If negated Then
            DescribeRequirement = "[Cannot be at: " & lid & "]"
        Else
            DescribeRequirement = "[Must be at: " & lid & "]"
        End If

    ElseIf modUtils.StartsWith(token, "MONEY:") Then
        DescribeRequirement = "[Requires $" & modUtils.StripPrefix(token, "MONEY:") & "]"

    Else
        DescribeRequirement = "[Locked]"
    End If
End Function
