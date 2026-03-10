Attribute VB_Name = "modUtils"
'===============================================================
' modUtils — Shared Utility Helpers
' Damned Moon VBA RPG Engine
'===============================================================
' Safe parsing, null guards, header lookups, weighted random,
' debug logging, token splitting — the little toolbox.
'===============================================================

Option Explicit

'===============================================================
' NULL / EMPTY GUARDS
'===============================================================

' Null-coalescing: returns fallback if value is Null, Empty, or ""
Public Function Nz(val As Variant, Optional fallback As Variant = "") As Variant
    If IsNull(val) Or IsEmpty(val) Then
        Nz = fallback
    ElseIf VarType(val) = vbString And CStr(val) = "" Then
        Nz = fallback
    Else
        Nz = val
    End If
End Function

' Safe CLng — returns fallback if not numeric
Public Function SafeLng(val As Variant, Optional fallback As Long = 0) As Long
    If IsNumeric(val) Then
        SafeLng = CLng(val)
    Else
        SafeLng = fallback
    End If
End Function

' Safe CDbl — returns fallback if not numeric
Public Function SafeDbl(val As Variant, Optional fallback As Double = 0#) As Double
    If IsNumeric(val) Then
        SafeDbl = CDbl(val)
    Else
        SafeDbl = fallback
    End If
End Function

' Safe CStr — never errors, always returns a string
Public Function SafeStr(val As Variant, Optional fallback As String = "") As String
    If IsNull(val) Or IsEmpty(val) Then
        SafeStr = fallback
    Else
        SafeStr = CStr(val & "")
        If SafeStr = "" Then SafeStr = fallback
    End If
End Function

' Safe CBool — returns False for non-boolean-like values
Public Function SafeBool(val As Variant, Optional fallback As Boolean = False) As Boolean
    If IsNull(val) Or IsEmpty(val) Then
        SafeBool = fallback
        Exit Function
    End If
    If VarType(val) = vbString Then
        Dim s As String
        s = UCase(Trim(CStr(val)))
        If s = "TRUE" Or s = "YES" Or s = "1" Then
            SafeBool = True
        Else
            SafeBool = False
        End If
    ElseIf IsNumeric(val) Then
        SafeBool = (CLng(val) <> 0)
    Else
        On Error Resume Next
        SafeBool = CBool(val)
        If Err.Number <> 0 Then SafeBool = fallback
        On Error GoTo 0
    End If
End Function

'===============================================================
' STRING / TOKEN HELPERS
'===============================================================

' Split that always returns a safe array (never errors on empty string)
Public Function SplitSafe(text As String, delim As String) As Variant
    If Len(text) = 0 Then
        SplitSafe = Array()
    Else
        SplitSafe = Split(text, delim)
    End If
End Function

' Trim all elements in a split array
Public Function SplitTrimmed(text As String, delim As String) As Variant
    Dim parts As Variant
    parts = SplitSafe(text, delim)

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim(CStr(parts(i)))
    Next i

    SplitTrimmed = parts
End Function

' Check if a string starts with a prefix
Public Function StartsWith(text As String, prefix As String) As Boolean
    StartsWith = (Left(text, Len(prefix)) = prefix)
End Function

' Extract the part after a prefix (e.g., "STAT:RAGE+5" with prefix "STAT:" returns "RAGE+5")
Public Function StripPrefix(text As String, prefix As String) As String
    If StartsWith(text, prefix) Then
        StripPrefix = Mid(text, Len(prefix) + 1)
    Else
        StripPrefix = text
    End If
End Function

'===============================================================
' TABLE / HEADER HELPERS
'===============================================================

' Get column index by header name in a ListObject table
Public Function GetColumnIndex(tbl As ListObject, colName As String) As Long
    If tbl Is Nothing Then
        GetColumnIndex = 0
        Exit Function
    End If

    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name = colName Then
            GetColumnIndex = col.Index
            Exit Function
        End If
    Next col
    GetColumnIndex = 0
End Function

' Get column index by header name on a plain worksheet (row 1 headers)
Public Function GetHeaderCol(ws As Worksheet, headerName As String) As Long
    If ws Is Nothing Then
        GetHeaderCol = 0
        Exit Function
    End If

    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If CStr(ws.Cells(1, c).Value) = headerName Then
            GetHeaderCol = c
            Exit Function
        End If
    Next c
    GetHeaderCol = 0
End Function

' Get the last used row in a column
Public Function GetLastRow(ws As Worksheet, Optional col As Long = 1) As Long
    If ws Is Nothing Then
        GetLastRow = 0
        Exit Function
    End If
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

'===============================================================
' MATH / RANDOM HELPERS
'===============================================================

' Clamp a value between min and max
Public Function Clamp(val As Long, minVal As Long, maxVal As Long) As Long
    If val < minVal Then
        Clamp = minVal
    ElseIf val > maxVal Then
        Clamp = maxVal
    Else
        Clamp = val
    End If
End Function

' Clamp a Double value
Public Function ClampDbl(val As Double, minVal As Double, maxVal As Double) As Double
    If val < minVal Then
        ClampDbl = minVal
    ElseIf val > maxVal Then
        ClampDbl = maxVal
    Else
        ClampDbl = val
    End If
End Function

' Random integer between low and high (inclusive)
Public Function RandBetween(low As Long, high As Long) As Long
    Randomize
    RandBetween = Int((high - low + 1) * Rnd + low)
End Function

' Weighted random pick from a Collection of weight values
' Returns the 1-based index of the chosen item
Public Function WeightedPick(weights As Collection) As Long
    If weights.Count = 0 Then
        WeightedPick = 0
        Exit Function
    End If

    Dim total As Double
    Dim i As Long
    For i = 1 To weights.Count
        total = total + CDbl(weights(i))
    Next i

    If total <= 0 Then
        WeightedPick = RandBetween(1, weights.Count)
        Exit Function
    End If

    Randomize
    Dim roll As Double
    roll = Rnd * total

    Dim cumulative As Double
    For i = 1 To weights.Count
        cumulative = cumulative + CDbl(weights(i))
        If roll <= cumulative Then
            WeightedPick = i
            Exit Function
        End If
    Next i

    WeightedPick = weights.Count
End Function

'===============================================================
' OPERATOR PARSING
'===============================================================

' Find the position of the first operator (+, -, =) in a string
' Returns 0 if none found
Public Function FindOperatorPos(expr As String) As Long
    Dim i As Long
    For i = 1 To Len(expr)
        Dim ch As String
        ch = Mid(expr, i, 1)
        If ch = "+" Or ch = "-" Or ch = "=" Then
            FindOperatorPos = i
            Exit Function
        End If
    Next i
    FindOperatorPos = 0
End Function

' Find the position of the first comparison operator (>, <, =, >=, <=)
' Returns 0 if none found
Public Function FindComparisonPos(expr As String) As Long
    Dim i As Long
    For i = 1 To Len(expr)
        Dim ch As String
        ch = Mid(expr, i, 1)
        If ch = ">" Or ch = "<" Or ch = "=" Then
            FindComparisonPos = i
            Exit Function
        End If
    Next i
    FindComparisonPos = 0
End Function

'===============================================================
' DEBUG LOGGING
'===============================================================

' Write debug message to Immediate window
Public Sub DebugLog(msg As String)
    If modConfig.DEBUG_MODE Then
        Debug.Print "[" & Format(Now, "hh:mm:ss") & "] " & msg
    End If
End Sub

' Always log (regardless of debug mode) — for errors
Public Sub ErrorLog(source As String, msg As String)
    Debug.Print "[ERROR " & Format(Now, "hh:mm:ss") & "] " & source & ": " & msg
End Sub
