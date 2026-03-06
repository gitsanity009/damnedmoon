Attribute VB_Name = "modEngine"
'===============================================================================
' modEngine — Choice Resolver & Scene Engine
' Blood Moon Protocol RPG Engine
'
' Orchestrates scene loading, choice resolution, condition validation,
' and effect application. This is the main game loop controller that
' ties modData (scene lookup) and modState (state management) together.
'
' Dependencies:
'   modData  — SceneRecord type, GetScene(), SceneExists()
'   modState — GetStat/SetStat/ModStat, GetFlag/SetFlag/HasFlag,
'              MoveToScene, TakeSnapshot/RestoreSnapshot, LogAction
'
' Effect/Condition JSON format (stored in SceneDB columns P-R):
'   Effects:    "HP:-10;Humanity:-5;SetFlag:MetAlpha;Rage:+20"
'   Conditions: "HasFlag:StoleSilverKnife;Humanity>=50;HP>0"
'===============================================================================
Option Explicit

' ── Engine state ──────────────────────────────────────────────────────────────
Private m_CurrentScene  As SceneRecord
Private m_SceneLoaded   As Boolean
Private m_UndoStack()   As String    ' Snapshot strings for rewind
Private m_UndoCount     As Long
Private Const MAX_UNDO  As Long = 50

' ── Result type for choice resolution ─────────────────────────────────────────
Public Type ChoiceResult
    Success       As Boolean    ' Was the choice valid and resolved?
    NextSceneID   As String     ' Scene we're moving to
    Message       As String     ' Feedback text (failure reason, effect summary)
    EffectsApplied As String    ' Description of effects that were applied
    ConditionFail As String     ' Which condition blocked the choice (if any)
End Type

' ══════════════════════════════════════════════════════════════════════════════
'  GAME LIFECYCLE
' ══════════════════════════════════════════════════════════════════════════════

Public Sub StartNewGame()
    ' Initialize a fresh game and load the title scene.
    ReDim m_UndoStack(0 To MAX_UNDO - 1)
    m_UndoCount = 0

    modState.InitNewGame
    LoadScene "TITLE"
End Sub

Public Sub ContinueGame()
    ' Resume from saved state.
    ReDim m_UndoStack(0 To MAX_UNDO - 1)
    m_UndoCount = 0

    modState.LoadState
    LoadScene modState.CurrentScene
End Sub

' ══════════════════════════════════════════════════════════════════════════════
'  SCENE LOADING
' ══════════════════════════════════════════════════════════════════════════════

Public Sub LoadScene(ByVal sceneID As String)
    ' Load a scene from the database, apply on-enter effects, update state.
    If Not modData.SceneExists(sceneID) Then
        modState.LogAction "ERROR", "Scene not found: " & sceneID
        Exit Sub
    End If

    ' Take undo snapshot BEFORE entering new scene
    PushUndo

    ' Load scene record
    m_CurrentScene = modData.GetScene(sceneID)
    m_SceneLoaded = True

    ' Apply on-enter effects (stat changes, flag sets, etc.)
    If Len(m_CurrentScene.OnEnterEffects) > 0 Then
        ApplyEffects m_CurrentScene.OnEnterEffects
    End If

    ' Auto-save on scene change
    modState.SaveState "Auto"

    modState.LogAction "LOAD_SCENE", sceneID & " (" & m_CurrentScene.SceneType & ")"

    ' Check for death/beast state after effects
    CheckFailState
End Sub

Public Function GetCurrentScene() As SceneRecord
    ' Returns the currently loaded scene record.
    If Not m_SceneLoaded Then
        ' Load from state if engine was just initialized
        Dim sid As String
        sid = modState.CurrentScene
        If Len(sid) > 0 And modData.SceneExists(sid) Then
            m_CurrentScene = modData.GetScene(sid)
            m_SceneLoaded = True
        End If
    End If
    GetCurrentScene = m_CurrentScene
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  CHOICE RESOLUTION
' ══════════════════════════════════════════════════════════════════════════════

Public Function ResolveChoice(ByVal choiceLetter As String) As ChoiceResult
    ' Main entry point: player picks "A" or "B" (or "C".."F" for future use).
    ' Validates conditions, applies effects, moves to next scene.
    Dim result As ChoiceResult
    result.Success = False

    If Not m_SceneLoaded Then
        result.Message = "No scene loaded."
        ResolveChoice = result
        Exit Function
    End If

    ' ── Get choice data based on letter ──
    Dim choiceLabel As String
    Dim choiceDesc As String
    Dim nextSceneID As String
    Dim condition As String

    Select Case UCase$(choiceLetter)
        Case "A"
            choiceLabel = m_CurrentScene.ChoiceA_Label
            choiceDesc = m_CurrentScene.ChoiceA_Desc
            nextSceneID = m_CurrentScene.ChoiceA_Next
            condition = m_CurrentScene.ConditionA
        Case "B"
            choiceLabel = m_CurrentScene.ChoiceB_Label
            choiceDesc = m_CurrentScene.ChoiceB_Desc
            nextSceneID = m_CurrentScene.ChoiceB_Next
            condition = m_CurrentScene.ConditionB
        Case Else
            result.Message = "Invalid choice: " & choiceLetter
            ResolveChoice = result
            Exit Function
    End Select

    ' ── Validate choice exists ──
    If Len(choiceLabel) = 0 Then
        result.Message = "Choice " & UCase$(choiceLetter) & " is not available in this scene."
        ResolveChoice = result
        Exit Function
    End If

    If Len(nextSceneID) = 0 Then
        result.Message = "Choice " & UCase$(choiceLetter) & " has no destination scene."
        ResolveChoice = result
        Exit Function
    End If

    ' ── Check conditions ──
    If Len(condition) > 0 Then
        Dim condResult As String
        condResult = EvaluateConditions(condition)
        If Len(condResult) > 0 Then
            result.ConditionFail = condResult
            result.Message = "Cannot choose " & UCase$(choiceLetter) & ": " & condResult
            modState.LogAction "CHOICE_BLOCKED", m_CurrentScene.SceneID & " -> " & _
                UCase$(choiceLetter) & ": " & condResult
            ResolveChoice = result
            Exit Function
        End If
    End If

    ' ── Validate destination exists ──
    If Not modData.SceneExists(nextSceneID) Then
        result.Message = "Destination scene not found: " & nextSceneID
        modState.LogAction "ERROR", "Dead link: " & m_CurrentScene.SceneID & " -> " & nextSceneID
        ResolveChoice = result
        Exit Function
    End If

    ' ── All checks passed — resolve the choice ──
    modState.MoveToScene nextSceneID, UCase$(choiceLetter)
    modState.LogAction "CHOICE", m_CurrentScene.SceneID & " -> " & UCase$(choiceLetter) & _
        " -> " & nextSceneID

    ' Load the destination scene (applies on-enter effects)
    LoadScene nextSceneID

    result.Success = True
    result.NextSceneID = nextSceneID
    result.Message = "Moved to " & nextSceneID
    ResolveChoice = result
End Function

Public Function IsChoiceAvailable(ByVal choiceLetter As String) As Boolean
    ' Quick check: does this choice exist AND are its conditions met?
    If Not m_SceneLoaded Then
        IsChoiceAvailable = False
        Exit Function
    End If

    Dim label As String
    Dim condition As String

    Select Case UCase$(choiceLetter)
        Case "A"
            label = m_CurrentScene.ChoiceA_Label
            condition = m_CurrentScene.ConditionA
        Case "B"
            label = m_CurrentScene.ChoiceB_Label
            condition = m_CurrentScene.ConditionB
        Case Else
            IsChoiceAvailable = False
            Exit Function
    End Select

    ' Must have a label
    If Len(label) = 0 Then
        IsChoiceAvailable = False
        Exit Function
    End If

    ' Must pass conditions (if any)
    If Len(condition) > 0 Then
        IsChoiceAvailable = (Len(EvaluateConditions(condition)) = 0)
    Else
        IsChoiceAvailable = True
    End If
End Function

Public Function GetAvailableChoices() As String
    ' Returns comma-separated list of available choice letters (e.g. "A,B").
    Dim choices As String
    If IsChoiceAvailable("A") Then choices = "A"
    If IsChoiceAvailable("B") Then
        If Len(choices) > 0 Then choices = choices & ","
        choices = choices & "B"
    End If
    GetAvailableChoices = choices
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  CONDITION EVALUATION
' ══════════════════════════════════════════════════════════════════════════════

Public Function EvaluateConditions(ByVal conditionStr As String) As String
    ' Evaluates a semicolon-delimited condition string.
    ' Returns "" if ALL conditions pass, or a description of the first failure.
    '
    ' Supported formats:
    '   "HasFlag:FlagName"        — flag must be truthy
    '   "NotFlag:FlagName"        — flag must be falsy or absent
    '   "StatName>=Value"         — numeric comparison (>=, <=, >, <, =)
    '   "HasItem:ItemName"        — shortcut for HasFlag:Has_ItemName
    '   "SceneVisited:SceneID"    — checks if scene is in history
    Dim conditions() As String
    conditions = Split(conditionStr, ";")

    Dim i As Long
    For i = LBound(conditions) To UBound(conditions)
        Dim cond As String
        cond = Trim$(conditions(i))
        If Len(cond) = 0 Then GoTo NextCond

        Dim failMsg As String
        failMsg = EvalSingleCondition(cond)
        If Len(failMsg) > 0 Then
            EvaluateConditions = failMsg
            Exit Function
        End If
NextCond:
    Next i

    EvaluateConditions = ""  ' All passed
End Function

Private Function EvalSingleCondition(ByVal cond As String) As String
    ' Evaluate one condition. Returns "" on pass, description on fail.
    EvalSingleCondition = ""

    ' ── HasFlag:Name ──
    If Left$(cond, 8) = "HasFlag:" Then
        Dim flagName As String
        flagName = Mid$(cond, 9)
        If Not modState.HasFlag(flagName) Then
            EvalSingleCondition = "Requires: " & flagName
        End If
        Exit Function
    End If

    ' ── NotFlag:Name ──
    If Left$(cond, 8) = "NotFlag:" Then
        Dim nfName As String
        nfName = Mid$(cond, 9)
        If modState.HasFlag(nfName) Then
            EvalSingleCondition = "Blocked by: " & nfName
        End If
        Exit Function
    End If

    ' ── HasItem:Name (shortcut for HasFlag:Has_ItemName) ──
    If Left$(cond, 8) = "HasItem:" Then
        Dim itemName As String
        itemName = Mid$(cond, 9)
        If Not modState.HasFlag("Has_" & itemName) Then
            EvalSingleCondition = "Requires item: " & itemName
        End If
        Exit Function
    End If

    ' ── SceneVisited:SceneID ──
    If Left$(cond, 13) = "SceneVisited:" Then
        Dim visitID As String
        visitID = Mid$(cond, 14)
        Dim prevScenes As Variant
        prevScenes = modState.GetPreviousScenes()
        Dim found As Boolean
        found = False
        If IsArray(prevScenes) Then
            Dim j As Long
            For j = LBound(prevScenes) To UBound(prevScenes)
                If StrComp(CStr(prevScenes(j)), visitID, vbTextCompare) = 0 Then
                    found = True
                    Exit For
                End If
            Next j
        End If
        If Not found Then
            EvalSingleCondition = "Must visit: " & visitID & " first"
        End If
        Exit Function
    End If

    ' ── Numeric comparison: StatName>=Value, StatName>Value, etc. ──
    Dim op As String
    Dim opPos As Long
    Dim ops As Variant
    ops = Array(">=", "<=", ">", "<", "=")

    Dim k As Long
    For k = LBound(ops) To UBound(ops)
        opPos = InStr(cond, CStr(ops(k)))
        If opPos > 0 Then
            op = CStr(ops(k))
            Exit For
        End If
    Next k

    If opPos > 0 Then
        Dim statName As String
        statName = Trim$(Left$(cond, opPos - 1))
        Dim threshold As Long
        threshold = CLng(Trim$(Mid$(cond, opPos + Len(op))))
        Dim actual As Long
        actual = CLng(modState.GetStat(statName))

        Dim passed As Boolean
        Select Case op
            Case ">=": passed = (actual >= threshold)
            Case "<=": passed = (actual <= threshold)
            Case ">":  passed = (actual > threshold)
            Case "<":  passed = (actual < threshold)
            Case "=":  passed = (actual = threshold)
        End Select

        If Not passed Then
            EvalSingleCondition = statName & " is " & actual & " (need " & op & threshold & ")"
        End If
        Exit Function
    End If

    ' Unknown condition format — pass with warning
    modState.LogAction "WARN", "Unknown condition format: " & cond
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  EFFECT APPLICATION
' ══════════════════════════════════════════════════════════════════════════════

Public Sub ApplyEffects(ByVal effectStr As String)
    ' Applies a semicolon-delimited effect string.
    '
    ' Supported formats:
    '   "HP:-10"              — modify stat by delta
    '   "HP=50"               — set stat to exact value
    '   "SetFlag:FlagName"    — set flag to True
    '   "ClearFlag:FlagName"  — remove flag
    '   "SetFlag:Name=Value"  — set flag to specific value
    '   "GiveItem:ItemName"   — shortcut for SetFlag:Has_ItemName=True
    '   "RemoveItem:ItemName" — shortcut for ClearFlag:Has_ItemName
    '   "TimeAdvance:30"      — advance TimeOfDay by N minutes
    '   "MoonAdvance:1"       — advance MoonPhase by N steps
    Dim effects() As String
    effects = Split(effectStr, ";")

    Dim summary As String
    Dim i As Long
    For i = LBound(effects) To UBound(effects)
        Dim eff As String
        eff = Trim$(effects(i))
        If Len(eff) = 0 Then GoTo NextEffect

        Dim desc As String
        desc = ApplySingleEffect(eff)
        If Len(desc) > 0 Then
            If Len(summary) > 0 Then summary = summary & ", "
            summary = summary & desc
        End If
NextEffect:
    Next i

    If Len(summary) > 0 Then
        modState.LogAction "EFFECTS", summary
    End If
End Sub

Private Function ApplySingleEffect(ByVal eff As String) As String
    ' Apply one effect. Returns description of what happened.
    ApplySingleEffect = ""

    ' ── SetFlag:Name or SetFlag:Name=Value ──
    If Left$(eff, 8) = "SetFlag:" Then
        Dim flagPart As String
        flagPart = Mid$(eff, 9)
        Dim eqPos As Long
        eqPos = InStr(flagPart, "=")
        If eqPos > 0 Then
            Dim fn As String
            fn = Left$(flagPart, eqPos - 1)
            Dim fv As String
            fv = Mid$(flagPart, eqPos + 1)
            modState.SetFlag fn, fv
            ApplySingleEffect = fn & "=" & fv
        Else
            modState.SetFlag flagPart, True
            ApplySingleEffect = flagPart & "=True"
        End If
        Exit Function
    End If

    ' ── ClearFlag:Name ──
    If Left$(eff, 10) = "ClearFlag:" Then
        Dim cfName As String
        cfName = Mid$(eff, 11)
        modState.ClearFlag cfName
        ApplySingleEffect = cfName & " cleared"
        Exit Function
    End If

    ' ── GiveItem:Name ──
    If Left$(eff, 9) = "GiveItem:" Then
        Dim giName As String
        giName = Mid$(eff, 10)
        modState.SetFlag "Has_" & giName, True
        Dim curItems As Long
        curItems = CLng(modState.GetStat("SilverItems"))
        If InStr(LCase$(giName), "silver") > 0 Then
            modState.ModStat "SilverItems", 1
        End If
        ApplySingleEffect = "Got: " & giName
        Exit Function
    End If

    ' ── RemoveItem:Name ──
    If Left$(eff, 11) = "RemoveItem:" Then
        Dim riName As String
        riName = Mid$(eff, 12)
        modState.ClearFlag "Has_" & riName
        ApplySingleEffect = "Lost: " & riName
        Exit Function
    End If

    ' ── TimeAdvance:N ──
    If Left$(eff, 12) = "TimeAdvance:" Then
        Dim mins As Long
        mins = CLng(Mid$(eff, 13))
        modState.ModStat "TimeOfDay", mins
        ApplySingleEffect = "Time +" & mins & "min"
        Exit Function
    End If

    ' ── MoonAdvance:N ──
    If Left$(eff, 12) = "MoonAdvance:" Then
        Dim steps As Long
        steps = CLng(Mid$(eff, 13))
        Dim curMoon As Long
        curMoon = CLng(modState.GetStat("MoonPhase"))
        Dim newMoon As Long
        newMoon = ((curMoon + steps - 1) Mod 8) + 1  ' Cycle 1-8
        modState.SetStat "MoonPhase", newMoon
        ApplySingleEffect = "Moon -> phase " & newMoon
        Exit Function
    End If

    ' ── Stat modification: "StatName:+N" or "StatName:-N" ──
    Dim colonPos As Long
    colonPos = InStr(eff, ":")
    If colonPos > 0 Then
        Dim sName As String
        sName = Left$(eff, colonPos - 1)
        Dim sVal As String
        sVal = Mid$(eff, colonPos + 1)

        ' Check it looks like a stat name (not a known prefix we missed)
        If Left$(sVal, 1) = "+" Or Left$(sVal, 1) = "-" Then
            Dim delta As Long
            delta = CLng(sVal)
            modState.ModStat sName, delta
            ApplySingleEffect = sName & " " & IIf(delta >= 0, "+", "") & delta
            Exit Function
        End If
    End If

    ' ── Stat set: "StatName=Value" ──
    Dim eqp As Long
    eqp = InStr(eff, "=")
    If eqp > 0 Then
        Dim setName As String
        setName = Trim$(Left$(eff, eqp - 1))
        Dim setVal As String
        setVal = Trim$(Mid$(eff, eqp + 1))
        If IsNumeric(setVal) Then
            modState.SetStat setName, CLng(setVal)
            ApplySingleEffect = setName & " = " & setVal
            Exit Function
        End If
    End If

    ' Unknown effect — log warning
    modState.LogAction "WARN", "Unknown effect: " & eff
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  UNDO / REWIND
' ══════════════════════════════════════════════════════════════════════════════

Private Sub PushUndo()
    ' Save current state snapshot for undo.
    If m_UndoCount >= MAX_UNDO Then
        ' Shift stack: discard oldest
        Dim i As Long
        For i = 0 To MAX_UNDO - 2
            m_UndoStack(i) = m_UndoStack(i + 1)
        Next i
        m_UndoCount = MAX_UNDO - 1
    End If
    m_UndoStack(m_UndoCount) = modState.TakeSnapshot()
    m_UndoCount = m_UndoCount + 1
End Sub

Public Function CanUndo() As Boolean
    CanUndo = (m_UndoCount > 0)
End Function

Public Sub Undo()
    ' Rewind to previous state snapshot.
    If m_UndoCount = 0 Then
        modState.LogAction "WARN", "Nothing to undo"
        Exit Sub
    End If

    m_UndoCount = m_UndoCount - 1
    modState.RestoreSnapshot m_UndoStack(m_UndoCount)

    ' Reload the scene record (don't re-apply on-enter effects)
    Dim sid As String
    sid = modState.CurrentScene
    If modData.SceneExists(sid) Then
        m_CurrentScene = modData.GetScene(sid)
        m_SceneLoaded = True
    End If

    modState.LogAction "UNDO", "Rewound to " & sid
End Sub

' ══════════════════════════════════════════════════════════════════════════════
'  NAVIGATION HELPERS
' ══════════════════════════════════════════════════════════════════════════════

Public Sub GoBack()
    ' Use modState's history stack to go back one scene.
    If Not modState.CanGoBack() Then
        modState.LogAction "WARN", "No history to go back to"
        Exit Sub
    End If

    Dim prevScene As String
    prevScene = modState.GoBack()
    If Len(prevScene) > 0 And modData.SceneExists(prevScene) Then
        m_CurrentScene = modData.GetScene(prevScene)
        m_SceneLoaded = True
    End If
End Sub

Public Sub ContinueScene()
    ' For transition scenes with a single "continue" link (Choice A only).
    If Not m_SceneLoaded Then Exit Sub

    If m_CurrentScene.SceneType = "transition" Or _
       (Len(m_CurrentScene.ChoiceA_Next) > 0 And Len(m_CurrentScene.ChoiceB_Label) = 0) Then
        Dim result As ChoiceResult
        result = ResolveChoice("A")
    End If
End Sub

' ══════════════════════════════════════════════════════════════════════════════
'  FAILURE STATE CHECKS
' ══════════════════════════════════════════════════════════════════════════════

Private Sub CheckFailState()
    ' Check if player has hit a death or beast-transformation state.
    If Not modState.IsAlive() Then
        modState.LogAction "DEATH", "HP reached 0"
        modState.SetFlag "IsDead", True
        ' Future: modUI will show death screen
    End If

    If Not modState.IsHuman() Then
        modState.LogAction "BEAST", "Humanity reached 0"
        modState.SetFlag "IsFullBeast", True
        ' Future: force transition to ENDING_BEAST
    End If
End Sub

Public Function IsGameOver() As Boolean
    ' Returns True if the game has ended (death, beast, or ending scene).
    IsGameOver = modState.HasFlag("IsDead") Or _
                 modState.HasFlag("IsFullBeast") Or _
                 (m_SceneLoaded And m_CurrentScene.SceneType = "ending")
End Function

' ══════════════════════════════════════════════════════════════════════════════
'  BUTTON CLICK HANDLERS (called from worksheet buttons)
' ══════════════════════════════════════════════════════════════════════════════

Public Sub OnChoiceA()
    ' Attached to Choice A button on game screen.
    Dim result As ChoiceResult
    result = ResolveChoice("A")
    If Not result.Success Then
        MsgBox result.Message, vbExclamation, "Blood Moon Protocol"
    End If
    ' Future: modUI.RenderScene will handle display update
End Sub

Public Sub OnChoiceB()
    ' Attached to Choice B button on game screen.
    Dim result As ChoiceResult
    result = ResolveChoice("B")
    If Not result.Success Then
        MsgBox result.Message, vbExclamation, "Blood Moon Protocol"
    End If
End Sub

Public Sub OnContinue()
    ' Attached to Continue button for transition/ending scenes.
    ContinueScene
End Sub

Public Sub OnBack()
    ' Attached to Back button on game screen.
    GoBack
End Sub

Public Sub OnUndo()
    ' Attached to Undo button on game screen.
    Undo
End Sub

Public Sub OnNewGame()
    ' Attached to New Game button.
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Start a new game? Current progress will be lost.", _
                    vbYesNo + vbQuestion, "Blood Moon Protocol")
    If answer = vbYes Then
        StartNewGame
    End If
End Sub
