Attribute VB_Name = "modData"
'===============================================================================
' modData — Scene Database Access Layer
' Blood Moon Protocol RPG Engine
'
' Reads scene data from the SceneDB sheet and provides lookup functions
' for the engine (modEngine), UI (modUI), and state manager (modState).
'
' SceneDB Sheet Layout (row 1 = headers, row 2 = descriptions, data from row 3):
'   A: SceneID        B: SceneTitle      C: StoryText       D: HP
'   E: Humanity       F: MoonPhase       G: ChoicePrompt    H: ChoiceA_Label
'   I: ChoiceA_Desc   J: ChoiceA_Next    K: ChoiceB_Label   L: ChoiceB_Desc
'   M: ChoiceB_Next   N: SceneType       O: Warning         P: OnEnterEffects
'   Q: ConditionA     R: ConditionB
'===============================================================================
Option Explicit

' ── Constants ─────────────────────────────────────────────────────────────────
Private Const SHEET_NAME   As String = "SceneDB"
Private Const DATA_ROW     As Long = 3          ' First data row (after header + desc)
Private Const COL_ID       As Long = 1          ' A
Private Const COL_TITLE    As Long = 2          ' B
Private Const COL_STORY    As Long = 3          ' C
Private Const COL_HP       As Long = 4          ' D
Private Const COL_HUM      As Long = 5          ' E
Private Const COL_MOON     As Long = 6          ' F
Private Const COL_PROMPT   As Long = 7          ' G
Private Const COL_A_LABEL  As Long = 8          ' H
Private Const COL_A_DESC   As Long = 9          ' I
Private Const COL_A_NEXT   As Long = 10         ' J
Private Const COL_B_LABEL  As Long = 11         ' K
Private Const COL_B_DESC   As Long = 12         ' L
Private Const COL_B_NEXT   As Long = 13         ' M
Private Const COL_TYPE     As Long = 14         ' N
Private Const COL_WARNING  As Long = 15         ' O
Private Const COL_EFFECTS  As Long = 16         ' P
Private Const COL_COND_A   As Long = 17         ' Q
Private Const COL_COND_B   As Long = 18         ' R
Private Const LAST_COL     As Long = 18

' ── Scene Data Type ───────────────────────────────────────────────────────────
Public Type SceneRecord
    SceneID        As String
    SceneTitle     As String
    StoryText      As String
    HP             As Long
    Humanity       As Long
    MoonPhase      As String
    ChoicePrompt   As String
    ChoiceA_Label  As String
    ChoiceA_Desc   As String
    ChoiceA_Next   As String
    ChoiceB_Label  As String
    ChoiceB_Desc   As String
    ChoiceB_Next   As String
    SceneType      As String    ' "choice", "transition", "ending", "title"
    Warning        As String
    OnEnterEffects As String    ' JSON-like string for future parsing
    ConditionA     As String    ' JSON-like condition for Choice A
    ConditionB     As String    ' JSON-like condition for Choice B
    RowIndex       As Long      ' Row number in SceneDB (for updates)
End Type

' ── Cached index: SceneID -> row number ───────────────────────────────────────
Private m_Index      As Object  ' Scripting.Dictionary
Private m_IndexBuilt As Boolean

' ── Private Helpers ───────────────────────────────────────────────────────────

Private Function GetDB() As Worksheet
    On Error Resume Next
    Set GetDB = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If GetDB Is Nothing Then
        Err.Raise vbObjectError + 1001, "modData.GetDB", _
            "SceneDB sheet not found. Run build_scene_db.py to create it."
    End If
End Function

Private Sub BuildIndex()
    ' Build a dictionary mapping SceneID -> row number for O(1) lookups.
    If m_IndexBuilt Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetDB()

    Set m_Index = CreateObject("Scripting.Dictionary")
    m_Index.CompareMode = vbTextCompare   ' Case-insensitive

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    Dim r As Long
    Dim sceneID As String
    For r = DATA_ROW To lastRow
        sceneID = Trim$(CStr(ws.Cells(r, COL_ID).Value))
        If Len(sceneID) > 0 Then
            m_Index(sceneID) = r
        End If
    Next r

    m_IndexBuilt = True
End Sub

Private Sub InvalidateIndex()
    ' Call this after any write to SceneDB so the cache is rebuilt on next read.
    m_IndexBuilt = False
    Set m_Index = Nothing
End Sub

' ── Public Read API ───────────────────────────────────────────────────────────

Public Function SceneExists(ByVal sceneID As String) As Boolean
    ' Returns True if the given SceneID exists in the database.
    BuildIndex
    SceneExists = m_Index.Exists(sceneID)
End Function

Public Function GetScene(ByVal sceneID As String) As SceneRecord
    ' Looks up a scene by ID and returns a populated SceneRecord.
    ' Raises an error if the scene is not found.
    BuildIndex

    If Not m_Index.Exists(sceneID) Then
        Err.Raise vbObjectError + 1002, "modData.GetScene", _
            "Scene not found: '" & sceneID & "'"
    End If

    Dim ws As Worksheet
    Set ws = GetDB()

    Dim r As Long
    r = m_Index(sceneID)

    Dim s As SceneRecord
    With s
        .SceneID        = SafeStr(ws.Cells(r, COL_ID).Value)
        .SceneTitle     = SafeStr(ws.Cells(r, COL_TITLE).Value)
        .StoryText      = SafeStr(ws.Cells(r, COL_STORY).Value)
        .HP             = SafeLng(ws.Cells(r, COL_HP).Value)
        .Humanity       = SafeLng(ws.Cells(r, COL_HUM).Value)
        .MoonPhase      = SafeStr(ws.Cells(r, COL_MOON).Value)
        .ChoicePrompt   = SafeStr(ws.Cells(r, COL_PROMPT).Value)
        .ChoiceA_Label  = SafeStr(ws.Cells(r, COL_A_LABEL).Value)
        .ChoiceA_Desc   = SafeStr(ws.Cells(r, COL_A_DESC).Value)
        .ChoiceA_Next   = SafeStr(ws.Cells(r, COL_A_NEXT).Value)
        .ChoiceB_Label  = SafeStr(ws.Cells(r, COL_B_LABEL).Value)
        .ChoiceB_Desc   = SafeStr(ws.Cells(r, COL_B_DESC).Value)
        .ChoiceB_Next   = SafeStr(ws.Cells(r, COL_B_NEXT).Value)
        .SceneType      = SafeStr(ws.Cells(r, COL_TYPE).Value)
        .Warning        = SafeStr(ws.Cells(r, COL_WARNING).Value)
        .OnEnterEffects = SafeStr(ws.Cells(r, COL_EFFECTS).Value)
        .ConditionA     = SafeStr(ws.Cells(r, COL_COND_A).Value)
        .ConditionB     = SafeStr(ws.Cells(r, COL_COND_B).Value)
        .RowIndex       = r
    End With

    GetScene = s
End Function

Public Function GetSceneField(ByVal sceneID As String, ByVal fieldName As String) As Variant
    ' Quick single-field lookup without loading the full record.
    ' fieldName is case-insensitive column header (e.g. "StoryText", "ChoiceA_Next").
    BuildIndex

    If Not m_Index.Exists(sceneID) Then
        GetSceneField = ""
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = GetDB()
    Dim r As Long
    r = m_Index(sceneID)

    Dim col As Long
    col = FieldNameToCol(fieldName)
    If col = 0 Then
        GetSceneField = ""
        Exit Function
    End If

    GetSceneField = ws.Cells(r, col).Value
End Function

Public Function GetAllSceneIDs() As Variant
    ' Returns a 1D array of all SceneIDs in order.
    BuildIndex
    GetAllSceneIDs = m_Index.Keys
End Function

Public Function GetSceneCount() As Long
    ' Returns the total number of scenes.
    BuildIndex
    GetSceneCount = m_Index.Count
End Function

Public Function GetChoiceCount(ByVal sceneID As String) As Long
    ' Returns 0, 1, or 2 depending on how many choices the scene has.
    Dim s As SceneRecord
    s = GetScene(sceneID)

    If Len(s.ChoiceB_Label) > 0 Then
        GetChoiceCount = 2
    ElseIf Len(s.ChoiceA_Label) > 0 Then
        GetChoiceCount = 1
    Else
        GetChoiceCount = 0
    End If
End Function

Public Function GetNextSceneID(ByVal sceneID As String, ByVal choiceLetter As String) As String
    ' Given a sceneID and "A" or "B", returns the next scene's ID.
    Dim s As SceneRecord
    s = GetScene(sceneID)

    Select Case UCase$(choiceLetter)
        Case "A": GetNextSceneID = s.ChoiceA_Next
        Case "B": GetNextSceneID = s.ChoiceB_Next
        Case Else: GetNextSceneID = ""
    End Select
End Function

' ── Public Write API ──────────────────────────────────────────────────────────

Public Sub SetSceneField(ByVal sceneID As String, ByVal fieldName As String, ByVal newValue As Variant)
    ' Update a single field for an existing scene.
    BuildIndex

    If Not m_Index.Exists(sceneID) Then
        Err.Raise vbObjectError + 1003, "modData.SetSceneField", _
            "Cannot update — scene not found: '" & sceneID & "'"
    End If

    Dim ws As Worksheet
    Set ws = GetDB()
    Dim r As Long
    r = m_Index(sceneID)

    Dim col As Long
    col = FieldNameToCol(fieldName)
    If col = 0 Then
        Err.Raise vbObjectError + 1004, "modData.SetSceneField", _
            "Unknown field: '" & fieldName & "'"
    End If

    ws.Cells(r, col).Value = newValue
End Sub

Public Sub AddScene(ByVal sceneID As String, Optional ByVal sceneTitle As String = "", _
                    Optional ByVal sceneType As String = "choice")
    ' Append a new blank scene row to SceneDB.
    BuildIndex

    If m_Index.Exists(sceneID) Then
        Err.Raise vbObjectError + 1005, "modData.AddScene", _
            "Scene already exists: '" & sceneID & "'"
    End If

    Dim ws As Worksheet
    Set ws = GetDB()

    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row + 1

    ws.Cells(newRow, COL_ID).Value = sceneID
    ws.Cells(newRow, COL_TITLE).Value = IIf(Len(sceneTitle) > 0, sceneTitle, sceneID)
    ws.Cells(newRow, COL_TYPE).Value = sceneType

    InvalidateIndex
End Sub

' ── Validation ────────────────────────────────────────────────────────────────

Public Function ValidateSceneDB() As String
    ' Checks for common issues: dead links, orphan scenes, missing fields.
    ' Returns a multi-line report string.
    BuildIndex

    Dim ws As Worksheet
    Set ws = GetDB()

    Dim report As String
    Dim issues As Long
    Dim allIDs As Variant
    allIDs = m_Index.Keys

    Dim reachable As Object
    Set reachable = CreateObject("Scripting.Dictionary")
    reachable.CompareMode = vbTextCompare

    Dim i As Long
    Dim s As SceneRecord

    For i = LBound(allIDs) To UBound(allIDs)
        s = GetScene(CStr(allIDs(i)))

        ' Check dead links
        If Len(s.ChoiceA_Next) > 0 Then
            reachable(s.ChoiceA_Next) = True
            If Not m_Index.Exists(s.ChoiceA_Next) Then
                report = report & "DEAD LINK: " & s.SceneID & " -> ChoiceA -> " & s.ChoiceA_Next & vbNewLine
                issues = issues + 1
            End If
        End If
        If Len(s.ChoiceB_Next) > 0 Then
            reachable(s.ChoiceB_Next) = True
            If Not m_Index.Exists(s.ChoiceB_Next) Then
                report = report & "DEAD LINK: " & s.SceneID & " -> ChoiceB -> " & s.ChoiceB_Next & vbNewLine
                issues = issues + 1
            End If
        End If

        ' Check missing story text
        If Len(s.StoryText) = 0 And s.SceneType <> "title" Then
            report = report & "EMPTY STORY: " & s.SceneID & vbNewLine
            issues = issues + 1
        End If

        ' Check choice scenes without choices
        If s.SceneType = "choice" And Len(s.ChoiceA_Label) = 0 Then
            report = report & "NO CHOICES: " & s.SceneID & " is type 'choice' but has no ChoiceA" & vbNewLine
            issues = issues + 1
        End If
    Next i

    ' Check orphan scenes (not reachable from any other scene, except TITLE)
    For i = LBound(allIDs) To UBound(allIDs)
        Dim sid As String
        sid = CStr(allIDs(i))
        If sid <> "TITLE" And Not reachable.Exists(sid) Then
            report = report & "ORPHAN: " & sid & " is not linked from any other scene" & vbNewLine
            issues = issues + 1
        End If
    Next i

    If issues = 0 Then
        ValidateSceneDB = "SceneDB validation passed. " & m_Index.Count & " scenes, 0 issues."
    Else
        ValidateSceneDB = "SceneDB validation: " & issues & " issue(s) found:" & vbNewLine & report
    End If
End Function

' ── Graph Export ──────────────────────────────────────────────────────────────

Public Sub ExportSceneGraph()
    ' Writes a simple adjacency list to a "SceneGraph" sheet for visualization.
    BuildIndex

    Dim wsGraph As Worksheet
    On Error Resume Next
    Set wsGraph = ThisWorkbook.Sheets("SceneGraph")
    On Error GoTo 0

    If wsGraph Is Nothing Then
        Set wsGraph = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsGraph.Name = "SceneGraph"
    Else
        wsGraph.Cells.Clear
    End If

    wsGraph.Cells(1, 1).Value = "From"
    wsGraph.Cells(1, 2).Value = "Choice"
    wsGraph.Cells(1, 3).Value = "To"
    wsGraph.Cells(1, 4).Value = "Type"

    Dim allIDs As Variant
    allIDs = m_Index.Keys

    Dim r As Long
    r = 2

    Dim i As Long
    Dim s As SceneRecord
    For i = LBound(allIDs) To UBound(allIDs)
        s = GetScene(CStr(allIDs(i)))

        If Len(s.ChoiceA_Next) > 0 Then
            wsGraph.Cells(r, 1).Value = s.SceneID
            wsGraph.Cells(r, 2).Value = "A"
            wsGraph.Cells(r, 3).Value = s.ChoiceA_Next
            wsGraph.Cells(r, 4).Value = s.SceneType
            r = r + 1
        End If
        If Len(s.ChoiceB_Next) > 0 Then
            wsGraph.Cells(r, 1).Value = s.SceneID
            wsGraph.Cells(r, 2).Value = "B"
            wsGraph.Cells(r, 3).Value = s.ChoiceB_Next
            wsGraph.Cells(r, 4).Value = s.SceneType
            r = r + 1
        End If
    Next i

    wsGraph.Columns.AutoFit
End Sub

' ── Utility Functions ─────────────────────────────────────────────────────────

Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

Private Function SafeLng(ByVal v As Variant) As Long
    If IsNull(v) Or IsEmpty(v) Then
        SafeLng = 0
    ElseIf IsNumeric(v) Then
        SafeLng = CLng(v)
    Else
        SafeLng = 0
    End If
End Function

Private Function FieldNameToCol(ByVal fieldName As String) As Long
    ' Map field name to column index. Returns 0 if unknown.
    Select Case LCase$(fieldName)
        Case "sceneid":        FieldNameToCol = COL_ID
        Case "scenetitle":     FieldNameToCol = COL_TITLE
        Case "storytext":      FieldNameToCol = COL_STORY
        Case "hp":             FieldNameToCol = COL_HP
        Case "humanity":       FieldNameToCol = COL_HUM
        Case "moonphase":      FieldNameToCol = COL_MOON
        Case "choiceprompt":   FieldNameToCol = COL_PROMPT
        Case "choicea_label":  FieldNameToCol = COL_A_LABEL
        Case "choicea_desc":   FieldNameToCol = COL_A_DESC
        Case "choicea_next":   FieldNameToCol = COL_A_NEXT
        Case "choiceb_label":  FieldNameToCol = COL_B_LABEL
        Case "choiceb_desc":   FieldNameToCol = COL_B_DESC
        Case "choiceb_next":   FieldNameToCol = COL_B_NEXT
        Case "scenetype":      FieldNameToCol = COL_TYPE
        Case "warning":        FieldNameToCol = COL_WARNING
        Case "onentereffects": FieldNameToCol = COL_EFFECTS
        Case "conditiona":     FieldNameToCol = COL_COND_A
        Case "conditionb":     FieldNameToCol = COL_COND_B
        Case Else:             FieldNameToCol = 0
    End Select
End Function
