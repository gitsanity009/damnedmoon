Attribute VB_Name = "modState"
'===============================================================================
' modState — Game State Manager
' Blood Moon Protocol RPG Engine
'
' Manages all runtime game state: current scene, stats, flags, inventory,
' scene history stack, and action log. Persists state to the hidden
' GameState sheet for save/load and debugging.
'
' GameState Sheet Layout:
'   "Stats"   section  (rows 3+):   StatName | Value
'   "Flags"   section  (rows 3+):   FlagName | Value  (col D-E)
'   "History" section  (rows 3+):   Index | SceneID | Choice | Timestamp  (col G-J)
'   "Log"     section  (rows 3+):   Index | Action | Detail | Timestamp   (col L-O)
'   "Meta"    section  (row 1):     CurrentScene | SaveSlot | SaveTimestamp (col Q-S)
'===============================================================================
Option Explicit

' ── Sheet & layout constants ──────────────────────────────────────────────────
Private Const WS_NAME       As String = "GameState"

' Stats block: columns A-B
Private Const STAT_COL_NAME As Long = 1   ' A
Private Const STAT_COL_VAL  As Long = 2   ' B
Private Const STAT_HDR_ROW  As Long = 2
Private Const STAT_DATA_ROW As Long = 3

' Flags block: columns D-E
Private Const FLAG_COL_NAME As Long = 4   ' D
Private Const FLAG_COL_VAL  As Long = 5   ' E
Private Const FLAG_HDR_ROW  As Long = 2
Private Const FLAG_DATA_ROW As Long = 3

' History stack: columns G-J
Private Const HIST_COL_IDX  As Long = 7   ' G
Private Const HIST_COL_SID  As Long = 8   ' H
Private Const HIST_COL_CHO  As Long = 9   ' I
Private Const HIST_COL_TS   As Long = 10  ' J
Private Const HIST_HDR_ROW  As Long = 2
Private Const HIST_DATA_ROW As Long = 3

' Action log: columns L-O
Private Const LOG_COL_IDX   As Long = 12  ' L
Private Const LOG_COL_ACT   As Long = 13  ' M
Private Const LOG_COL_DET   As Long = 14  ' N
Private Const LOG_COL_TS    As Long = 15  ' O
Private Const LOG_HDR_ROW   As Long = 2
Private Const LOG_DATA_ROW  As Long = 3

' Meta: columns Q-S, row 1
Private Const META_COL_SCENE As Long = 17 ' Q
Private Const META_COL_SLOT  As Long = 18 ' R
Private Const META_COL_SAVTS As Long = 19 ' S
Private Const META_ROW       As Long = 1

' ── Default starting stats ────────────────────────────────────────────────────
Private Const DEFAULT_STATS As String = _
    "HP:100|MaxHP:100|Humanity:100|MaxHumanity:100|Rage:0|MaxRage:100|" & _
    "Hunger:0|MaxHunger:100|ScentLevel:0|Suspicion:0|SilverItems:0|" & _
    "MoonPhase:5|TimeOfDay:347|TransformStage:0"

Private Const DEFAULT_SCENE As String = "TITLE"

' ── In-memory state cache ─────────────────────────────────────────────────────
Private m_Stats     As Object  ' Scripting.Dictionary  StatName -> Long/Double
Private m_Flags     As Object  ' Scripting.Dictionary  FlagName -> Variant
Private m_History() As String  ' Stack of "SceneID|Choice" entries
Private m_HistCount As Long
Private m_LogCount  As Long
Private m_CurScene  As String
Private m_Loaded    As Boolean

' ══════════════════════════════════════════════════════════════════════════════
'  INITIALIZATION
' ══════════════════════════════════════════════════════════════════════════════

Public Sub InitNewGame()
    ' Start a fresh game with default stats and flags.
    Set m_Stats = CreateObject("Scripting.Dictionary")
    m_Stats.CompareMode = vbTextCompare
    Set m_Flags = CreateObject("Scripting.Dictionary")
    m_Flags.CompareMode = vbTextCompare

    ' Parse default stats
    Dim pairs() As String
    pairs = Split(DEFAULT_STATS, "|")
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        If Len(pairs(i)) > 0 Then
            Dim kv() As String
            kv = Split(pairs(i), ":")
            If UBound(kv) >= 1 Then
                If IsNumeric(kv(1)) Then
                    m_Stats(kv(0)) = CLng(kv(1))
                Else
                    m_Stats(kv(0)) = kv(1)
                End If
            End If
        End If
    Next i

    ' Reset history
    ReDim m_History(0 To 99)
    m_HistCount = 0
    m_LogCount = 0

    m_CurScene = DEFAULT_SCENE
    m_Loaded = True

    ' Write to sheet
    EnsureSheet
    ClearSheet
    FlushAll
    LogAction "GAME_START", "New game initialized"
End Sub

' ══════════════════════════════════════════════════════════════════════════════
'  STAT ACCESS  (numeric values: HP, Humanity, Rage, Hunger, etc.)
' ══════════════════════════════════════════════════════════════════════════════

Public Function GetStat(ByVal statName As String) As Variant
    ' Returns stat value. Returns 0 if stat doesn't exist.
    EnsureLoaded
    If m_Stats.Exists(statName) Then
        GetStat = m_Stats(statName)
    Else
        GetStat = 0
    End If
End Function

Public Sub SetStat(ByVal statName As String, ByVal newValue As Variant)
    ' Set stat to exact value. Clamps to 0..Max if a Max exists.
    EnsureLoaded
    Dim clamped As Variant
    clamped = ClampStat(statName, newValue)
    m_Stats(statName) = clamped
    FlushStats
    LogAction "SET_STAT", statName & " = " & CStr(clamped)
End Sub

Public Sub ModStat(ByVal statName As String, ByVal delta As Long)
    ' Add delta to stat (can be negative). Clamps automatically.
    EnsureLoaded
    Dim cur As Long
    If m_Stats.Exists(statName) Then
        cur = CLng(m_Stats(statName))
    End If
    SetStat statName, cur + delta
End Sub

Public Function GetDerivedStat(ByVal derivedName As String) As Long
    ' Compute derived stats from base stats.
    EnsureLoaded
    Select Case LCase$(derivedName)
        Case "control"
            ' Control = Humanity - Rage (how well you resist transformation)
            GetDerivedStat = CLng(GetStat("Humanity")) - CLng(GetStat("Rage"))
        Case "threat"
            ' Threat = Rage + ScentLevel (how much danger you attract)
            GetDerivedStat = CLng(GetStat("Rage")) + CLng(GetStat("ScentLevel"))
        Case "transformrisk"
            ' TransformRisk = Rage + Hunger + (MoonPhase * 10) - Humanity
            GetDerivedStat = CLng(GetStat("Rage")) + CLng(GetStat("Hunger")) + _
                             (CLng(GetStat("MoonPhase")) * 10) - CLng(GetStat("Humanity"))
        Case Else
            GetDerivedStat = 0
    End Select
End Function

Public Function IsAlive() As Boolean
    ' Returns False if HP <= 0 (death state).
    EnsureLoaded
    IsAlive = (CLng(GetStat("HP")) > 0)
End Function

Public Function IsHuman() As Boolean
    ' Returns False if Humanity <= 0 (beast state / game over).
    EnsureLoaded
    IsHuman = (CLng(GetStat("Humanity")) > 0)
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  FLAG ACCESS  (boolean or string values: MetHunter, StoleSilverKnife, etc.)
' ══════════════════════════════════════════════════════════════════════════════

Public Function GetFlag(ByVal flagName As String) As Variant
    ' Returns flag value. Returns False if flag doesn't exist.
    EnsureLoaded
    If m_Flags.Exists(flagName) Then
        GetFlag = m_Flags(flagName)
    Else
        GetFlag = False
    End If
End Function

Public Sub SetFlag(ByVal flagName As String, ByVal newValue As Variant)
    ' Set any flag to any value (boolean, string, number).
    EnsureLoaded
    m_Flags(flagName) = newValue
    FlushFlags
    LogAction "SET_FLAG", flagName & " = " & CStr(newValue)
End Sub

Public Function HasFlag(ByVal flagName As String) As Boolean
    ' Returns True if flag exists AND is truthy (True, non-zero, non-empty).
    EnsureLoaded
    If Not m_Flags.Exists(flagName) Then
        HasFlag = False
        Exit Function
    End If
    Dim v As Variant
    v = m_Flags(flagName)
    Select Case VarType(v)
        Case vbBoolean:  HasFlag = CBool(v)
        Case vbString:   HasFlag = (Len(CStr(v)) > 0 And LCase$(CStr(v)) <> "false")
        Case vbLong, vbInteger, vbDouble: HasFlag = (CDbl(v) <> 0)
        Case Else:       HasFlag = False
    End Select
End Function

Public Sub ClearFlag(ByVal flagName As String)
    ' Remove a flag entirely.
    EnsureLoaded
    If m_Flags.Exists(flagName) Then
        m_Flags.Remove flagName
        FlushFlags
        LogAction "CLEAR_FLAG", flagName
    End If
End Sub

' ══════════════════════════════════════════════════════════════════════════════
'  SCENE NAVIGATION & HISTORY
' ══════════════════════════════════════════════════════════════════════════════

Public Property Get CurrentScene() As String
    EnsureLoaded
    CurrentScene = m_CurScene
End Property

Public Sub MoveToScene(ByVal newSceneID As String, Optional ByVal choiceMade As String = "")
    ' Navigate to a new scene. Pushes current scene onto history stack.
    EnsureLoaded

    ' Push current to history
    If Len(m_CurScene) > 0 Then
        PushHistory m_CurScene, choiceMade
    End If

    m_CurScene = newSceneID
    FlushMeta
    LogAction "MOVE", newSceneID & IIf(Len(choiceMade) > 0, " (choice: " & choiceMade & ")", "")
End Sub

Public Function CanGoBack() As Boolean
    ' Returns True if there's history to go back to.
    EnsureLoaded
    CanGoBack = (m_HistCount > 0)
End Function

Public Function GoBack() As String
    ' Pop the last scene from history and return to it.
    ' Returns the scene ID we went back to, or "" if no history.
    EnsureLoaded
    If m_HistCount = 0 Then
        GoBack = ""
        Exit Function
    End If

    m_HistCount = m_HistCount - 1
    Dim entry As String
    entry = m_History(m_HistCount)

    ' Parse "SceneID|Choice"
    Dim parts() As String
    parts = Split(entry, "|")
    m_CurScene = parts(0)

    FlushMeta
    FlushHistory
    LogAction "GO_BACK", "Returned to " & m_CurScene

    GoBack = m_CurScene
End Function

Public Function GetHistoryDepth() As Long
    EnsureLoaded
    GetHistoryDepth = m_HistCount
End Function

Public Function GetPreviousScenes() As Variant
    ' Returns array of scene IDs visited (most recent last).
    EnsureLoaded
    If m_HistCount = 0 Then
        GetPreviousScenes = Array()
        Exit Function
    End If
    Dim result() As String
    ReDim result(0 To m_HistCount - 1)
    Dim i As Long
    For i = 0 To m_HistCount - 1
        Dim parts() As String
        parts = Split(m_History(i), "|")
        result(i) = parts(0)
    Next i
    GetPreviousScenes = result
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  ACTION LOG  (for replay / debug)
' ══════════════════════════════════════════════════════════════════════════════

Public Sub LogAction(ByVal action As String, ByVal detail As String)
    ' Append an entry to the action log.
    Dim ws As Worksheet
    Set ws = GetWS()

    m_LogCount = m_LogCount + 1
    Dim r As Long
    r = LOG_DATA_ROW + m_LogCount - 1

    ws.Cells(r, LOG_COL_IDX).Value = m_LogCount
    ws.Cells(r, LOG_COL_ACT).Value = action
    ws.Cells(r, LOG_COL_DET).Value = detail
    ws.Cells(r, LOG_COL_TS).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Function GetLogCount() As Long
    GetLogCount = m_LogCount
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  SAVE / LOAD
' ══════════════════════════════════════════════════════════════════════════════

Public Sub SaveState(Optional ByVal slotName As String = "Auto")
    ' Persist current in-memory state to the GameState sheet.
    ' The sheet IS the save — this just timestamps it.
    EnsureLoaded
    FlushAll

    Dim ws As Worksheet
    Set ws = GetWS()
    ws.Cells(META_ROW, META_COL_SLOT).Value = slotName
    ws.Cells(META_ROW, META_COL_SAVTS).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    LogAction "SAVE", "Slot: " & slotName
End Sub

Public Sub LoadState()
    ' Read state from GameState sheet into memory.
    Dim ws As Worksheet
    Set ws = GetWS()

    ' ── Load meta ──
    m_CurScene = SafeStr(ws.Cells(META_ROW, META_COL_SCENE).Value)
    If Len(m_CurScene) = 0 Then m_CurScene = DEFAULT_SCENE

    ' ── Load stats ──
    Set m_Stats = CreateObject("Scripting.Dictionary")
    m_Stats.CompareMode = vbTextCompare
    Dim r As Long
    r = STAT_DATA_ROW
    Do While Len(SafeStr(ws.Cells(r, STAT_COL_NAME).Value)) > 0
        m_Stats(SafeStr(ws.Cells(r, STAT_COL_NAME).Value)) = ws.Cells(r, STAT_COL_VAL).Value
        r = r + 1
    Loop

    ' ── Load flags ──
    Set m_Flags = CreateObject("Scripting.Dictionary")
    m_Flags.CompareMode = vbTextCompare
    r = FLAG_DATA_ROW
    Do While Len(SafeStr(ws.Cells(r, FLAG_COL_NAME).Value)) > 0
        m_Flags(SafeStr(ws.Cells(r, FLAG_COL_NAME).Value)) = ws.Cells(r, FLAG_COL_VAL).Value
        r = r + 1
    Loop

    ' ── Load history ──
    ReDim m_History(0 To 99)
    m_HistCount = 0
    r = HIST_DATA_ROW
    Do While Len(SafeStr(ws.Cells(r, HIST_COL_SID).Value)) > 0
        Dim sid As String
        sid = SafeStr(ws.Cells(r, HIST_COL_SID).Value)
        Dim cho As String
        cho = SafeStr(ws.Cells(r, HIST_COL_CHO).Value)
        m_History(m_HistCount) = sid & "|" & cho
        m_HistCount = m_HistCount + 1
        If m_HistCount > UBound(m_History) Then
            ReDim Preserve m_History(0 To UBound(m_History) + 100)
        End If
        r = r + 1
    Loop

    ' ── Count log entries ──
    m_LogCount = 0
    r = LOG_DATA_ROW
    Do While Len(SafeStr(ws.Cells(r, LOG_COL_ACT).Value)) > 0
        m_LogCount = m_LogCount + 1
        r = r + 1
    Loop

    m_Loaded = True
    LogAction "LOAD", "State loaded from sheet"
End Sub

Public Sub ResetGame()
    ' Wipe all state and start fresh.
    LogAction "RESET", "Game reset requested"
    InitNewGame
End Sub

' ══════════════════════════════════════════════════════════════════════════════
'  SNAPSHOT (for undo / rewind)
' ══════════════════════════════════════════════════════════════════════════════

Public Function TakeSnapshot() As String
    ' Serializes current state to a pipe-delimited string for stack-based undo.
    ' Format: "SCENE:id|STAT:name=val|...|FLAG:name=val|..."
    EnsureLoaded
    Dim parts() As String
    ReDim parts(0 To m_Stats.Count + m_Flags.Count)

    Dim idx As Long
    idx = 0
    parts(idx) = "SCENE:" & m_CurScene
    idx = idx + 1

    Dim k As Variant
    For Each k In m_Stats.Keys
        parts(idx) = "STAT:" & CStr(k) & "=" & CStr(m_Stats(k))
        idx = idx + 1
    Next k
    For Each k In m_Flags.Keys
        parts(idx) = "FLAG:" & CStr(k) & "=" & CStr(m_Flags(k))
        idx = idx + 1
    Next k

    ReDim Preserve parts(0 To idx - 1)
    TakeSnapshot = Join(parts, "|")
End Function

Public Sub RestoreSnapshot(ByVal snapshot As String)
    ' Restores state from a snapshot string (created by TakeSnapshot).
    EnsureLoaded
    Dim tokens() As String
    tokens = Split(snapshot, "|")

    Set m_Stats = CreateObject("Scripting.Dictionary")
    m_Stats.CompareMode = vbTextCompare
    Set m_Flags = CreateObject("Scripting.Dictionary")
    m_Flags.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        Dim token As String
        token = tokens(i)
        If Left$(token, 6) = "SCENE:" Then
            m_CurScene = Mid$(token, 7)
        ElseIf Left$(token, 5) = "STAT:" Then
            Dim skv As String
            skv = Mid$(token, 6)
            Dim sp() As String
            sp = Split(skv, "=")
            If UBound(sp) >= 1 Then
                If IsNumeric(sp(1)) Then
                    m_Stats(sp(0)) = CLng(sp(1))
                Else
                    m_Stats(sp(0)) = sp(1)
                End If
            End If
        ElseIf Left$(token, 5) = "FLAG:" Then
            Dim fkv As String
            fkv = Mid$(token, 6)
            Dim fp() As String
            fp = Split(fkv, "=")
            If UBound(fp) >= 1 Then
                m_Flags(fp(0)) = fp(1)
            End If
        End If
    Next i

    FlushAll
    LogAction "RESTORE", "Snapshot restored to " & m_CurScene
End Sub

' ══════════════════════════════════════════════════════════════════════════════
'  STATE DUMP (debug helper)
' ══════════════════════════════════════════════════════════════════════════════

Public Function DumpState() As String
    ' Returns a human-readable dump of current state for debugging.
    EnsureLoaded
    Dim s As String
    s = "=== GAME STATE ===" & vbNewLine
    s = s & "Scene: " & m_CurScene & vbNewLine
    s = s & "History depth: " & m_HistCount & vbNewLine
    s = s & vbNewLine & "── Stats ──" & vbNewLine

    Dim k As Variant
    For Each k In m_Stats.Keys
        s = s & "  " & CStr(k) & " = " & CStr(m_Stats(k)) & vbNewLine
    Next k

    s = s & vbNewLine & "── Flags ──" & vbNewLine
    For Each k In m_Flags.Keys
        s = s & "  " & CStr(k) & " = " & CStr(m_Flags(k)) & vbNewLine
    Next k

    s = s & vbNewLine & "── Derived ──" & vbNewLine
    s = s & "  Control = " & GetDerivedStat("Control") & vbNewLine
    s = s & "  Threat = " & GetDerivedStat("Threat") & vbNewLine
    s = s & "  TransformRisk = " & GetDerivedStat("TransformRisk") & vbNewLine
    s = s & vbNewLine & "Alive: " & IsAlive() & "  Human: " & IsHuman() & vbNewLine
    s = s & "Log entries: " & m_LogCount & vbNewLine

    DumpState = s
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  PRIVATE — Sheet management
' ══════════════════════════════════════════════════════════════════════════════

Private Function GetWS() As Worksheet
    On Error Resume Next
    Set GetWS = ThisWorkbook.Sheets(WS_NAME)
    On Error GoTo 0
    If GetWS Is Nothing Then
        EnsureSheet
        Set GetWS = ThisWorkbook.Sheets(WS_NAME)
    End If
End Function

Private Sub EnsureSheet()
    ' Create the GameState sheet if it doesn't exist, hidden.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = WS_NAME
    End If

    ws.Visible = xlSheetVeryHidden

    ' Write section headers
    ws.Cells(STAT_HDR_ROW, STAT_COL_NAME).Value = "StatName"
    ws.Cells(STAT_HDR_ROW, STAT_COL_VAL).Value = "Value"
    ws.Cells(FLAG_HDR_ROW, FLAG_COL_NAME).Value = "FlagName"
    ws.Cells(FLAG_HDR_ROW, FLAG_COL_VAL).Value = "Value"
    ws.Cells(HIST_HDR_ROW, HIST_COL_IDX).Value = "#"
    ws.Cells(HIST_HDR_ROW, HIST_COL_SID).Value = "SceneID"
    ws.Cells(HIST_HDR_ROW, HIST_COL_CHO).Value = "Choice"
    ws.Cells(HIST_HDR_ROW, HIST_COL_TS).Value = "Timestamp"
    ws.Cells(LOG_HDR_ROW, LOG_COL_IDX).Value = "#"
    ws.Cells(LOG_HDR_ROW, LOG_COL_ACT).Value = "Action"
    ws.Cells(LOG_HDR_ROW, LOG_COL_DET).Value = "Detail"
    ws.Cells(LOG_HDR_ROW, LOG_COL_TS).Value = "Timestamp"

    ' Section labels in row 1
    ws.Cells(1, STAT_COL_NAME).Value = "STATS"
    ws.Cells(1, FLAG_COL_NAME).Value = "FLAGS"
    ws.Cells(1, HIST_COL_IDX).Value = "HISTORY"
    ws.Cells(1, LOG_COL_IDX).Value = "ACTION LOG"
    ws.Cells(1, META_COL_SCENE).Value = "META"
End Sub

Private Sub ClearSheet()
    ' Clear all data rows (preserve headers).
    Dim ws As Worksheet
    Set ws = GetWS()

    ' Clear stat data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, STAT_COL_NAME).End(xlUp).Row
    If lastRow >= STAT_DATA_ROW Then
        ws.Range(ws.Cells(STAT_DATA_ROW, STAT_COL_NAME), ws.Cells(lastRow, STAT_COL_VAL)).ClearContents
    End If

    ' Clear flag data
    lastRow = ws.Cells(ws.Rows.Count, FLAG_COL_NAME).End(xlUp).Row
    If lastRow >= FLAG_DATA_ROW Then
        ws.Range(ws.Cells(FLAG_DATA_ROW, FLAG_COL_NAME), ws.Cells(lastRow, FLAG_COL_VAL)).ClearContents
    End If

    ' Clear history
    lastRow = ws.Cells(ws.Rows.Count, HIST_COL_SID).End(xlUp).Row
    If lastRow >= HIST_DATA_ROW Then
        ws.Range(ws.Cells(HIST_DATA_ROW, HIST_COL_IDX), ws.Cells(lastRow, HIST_COL_TS)).ClearContents
    End If

    ' Clear log
    lastRow = ws.Cells(ws.Rows.Count, LOG_COL_ACT).End(xlUp).Row
    If lastRow >= LOG_DATA_ROW Then
        ws.Range(ws.Cells(LOG_DATA_ROW, LOG_COL_IDX), ws.Cells(lastRow, LOG_COL_TS)).ClearContents
    End If
End Sub

Private Sub EnsureLoaded()
    ' Load from sheet if not yet loaded.
    If Not m_Loaded Then
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_NAME)
        On Error GoTo 0
        If ws Is Nothing Then
            InitNewGame
        Else
            LoadState
        End If
    End If
End Sub

' ── Flush helpers (write in-memory state to sheet) ────────────────────────────

Private Sub FlushAll()
    FlushStats
    FlushFlags
    FlushHistory
    FlushMeta
End Sub

Private Sub FlushStats()
    Dim ws As Worksheet
    Set ws = GetWS()

    ' Clear existing
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, STAT_COL_NAME).End(xlUp).Row
    If lastRow >= STAT_DATA_ROW Then
        ws.Range(ws.Cells(STAT_DATA_ROW, STAT_COL_NAME), ws.Cells(lastRow, STAT_COL_VAL)).ClearContents
    End If

    ' Write current
    Dim r As Long
    r = STAT_DATA_ROW
    Dim k As Variant
    For Each k In m_Stats.Keys
        ws.Cells(r, STAT_COL_NAME).Value = CStr(k)
        ws.Cells(r, STAT_COL_VAL).Value = m_Stats(k)
        r = r + 1
    Next k
End Sub

Private Sub FlushFlags()
    Dim ws As Worksheet
    Set ws = GetWS()

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, FLAG_COL_NAME).End(xlUp).Row
    If lastRow >= FLAG_DATA_ROW Then
        ws.Range(ws.Cells(FLAG_DATA_ROW, FLAG_COL_NAME), ws.Cells(lastRow, FLAG_COL_VAL)).ClearContents
    End If

    Dim r As Long
    r = FLAG_DATA_ROW
    Dim k As Variant
    For Each k In m_Flags.Keys
        ws.Cells(r, FLAG_COL_NAME).Value = CStr(k)
        ws.Cells(r, FLAG_COL_VAL).Value = m_Flags(k)
        r = r + 1
    Next k
End Sub

Private Sub FlushHistory()
    Dim ws As Worksheet
    Set ws = GetWS()

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, HIST_COL_SID).End(xlUp).Row
    If lastRow >= HIST_DATA_ROW Then
        ws.Range(ws.Cells(HIST_DATA_ROW, HIST_COL_IDX), ws.Cells(lastRow, HIST_COL_TS)).ClearContents
    End If

    Dim r As Long
    Dim i As Long
    For i = 0 To m_HistCount - 1
        r = HIST_DATA_ROW + i
        Dim parts() As String
        parts = Split(m_History(i), "|")
        ws.Cells(r, HIST_COL_IDX).Value = i + 1
        ws.Cells(r, HIST_COL_SID).Value = parts(0)
        If UBound(parts) >= 1 Then ws.Cells(r, HIST_COL_CHO).Value = parts(1)
        ws.Cells(r, HIST_COL_TS).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Next i
End Sub

Private Sub FlushMeta()
    Dim ws As Worksheet
    Set ws = GetWS()
    ws.Cells(META_ROW, META_COL_SCENE).Value = m_CurScene
End Sub

' ── History stack helpers ─────────────────────────────────────────────────────

Private Sub PushHistory(ByVal sceneID As String, ByVal choice As String)
    If m_HistCount > UBound(m_History) Then
        ReDim Preserve m_History(0 To UBound(m_History) + 100)
    End If
    m_History(m_HistCount) = sceneID & "|" & choice
    m_HistCount = m_HistCount + 1
    FlushHistory
End Sub

' ── Stat clamping ─────────────────────────────────────────────────────────────

Private Function ClampStat(ByVal statName As String, ByVal rawValue As Variant) As Variant
    ' Clamp stat between 0 and its Max counterpart (if one exists).
    If Not IsNumeric(rawValue) Then
        ClampStat = rawValue
        Exit Function
    End If

    Dim v As Long
    v = CLng(rawValue)

    ' Floor at 0
    If v < 0 Then v = 0

    ' Check for Max cap
    Dim maxKey As String
    maxKey = "Max" & statName
    If m_Stats.Exists(maxKey) Then
        Dim cap As Long
        cap = CLng(m_Stats(maxKey))
        If v > cap Then v = cap
    End If

    ClampStat = v
End Function

' ── Utility ───────────────────────────────────────────────────────────────────

Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function
