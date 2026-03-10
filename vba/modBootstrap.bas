Attribute VB_Name = "modBootstrap"
'===============================================================
' modBootstrap — Game Initialization & Entry Point
' Damned Moon VBA RPG Engine
'===============================================================
' Starts the game, initializes state, builds caches, wires
' everything together. The clean front door.
'===============================================================

Option Explicit

' ── INIT STATE ──
Private mGameInitialized As Boolean

'===============================================================
' NEW GAME — Full fresh start
'===============================================================
Public Sub StartNewGame()
    If MsgBox("Start a new game? All progress will be lost.", _
              vbYesNo + vbExclamation, "NEW GAME") = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ' Initialize engine
    InitializeGame

    ' Reset all state
    modState.ResetGameState

    ' Get starting scene from config (or use default)
    Dim startScene As String
    startScene = modConfig.GetConfigValue("StartingScene", modConfig.DEFAULT_START_SCENE)

    ' Set starting location
    Dim startLocation As String
    startLocation = modConfig.GetConfigValue("StartingLocation", modConfig.DEFAULT_START_LOCATION)
    modState.SetCurrentLocation startLocation

    ' Load the first scene
    modSceneEngine.LoadScene startScene

    Application.ScreenUpdating = True

    modUtils.DebugLog "modBootstrap.StartNewGame: started at " & startScene
End Sub

'===============================================================
' CONTINUE GAME — Resume from autosave or last state
'===============================================================
Public Sub ContinueGame()
    Application.ScreenUpdating = False

    ' Initialize engine
    InitializeGame

    ' Check if there's a current scene saved
    Dim currentScene As String
    currentScene = modState.GetCurrentScene()

    If currentScene = "" Then
        ' No save state — fall through to new game
        Application.ScreenUpdating = True
        MsgBox "No saved game found. Starting new game.", vbInformation, "CONTINUE"
        StartNewGame
        Exit Sub
    End If

    ' Reload the current scene
    modSceneEngine.LoadScene currentScene
    modUtils.DebugLog "modBootstrap.ContinueGame: resuming at " & currentScene

    Application.ScreenUpdating = True
End Sub

'===============================================================
' INITIALIZE GAME — Core engine setup
'===============================================================
Public Sub InitializeGame()
    If mGameInitialized Then
        modUtils.DebugLog "modBootstrap.InitializeGame: already initialized, rebuilding caches"
    End If

    modUtils.DebugLog "modBootstrap.InitializeGame: starting"

    ' 1. Build all data caches
    modData.BuildCaches

    ' 2. Validate critical sheets exist
    ValidateCriticalSheets

    ' 3. Mark initialized
    mGameInitialized = True

    modUtils.DebugLog "modBootstrap.InitializeGame: complete"
End Sub

'===============================================================
' SETUP GAME — First-time setup (create buttons, etc.)
'===============================================================
Public Sub SetupGame()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then
        MsgBox "Game sheet not found!", vbCritical, "SETUP ERROR"
        Exit Sub
    End If

    ' Delete old choice buttons
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, Len(modConfig.BTN_PREFIX)) = modConfig.BTN_PREFIX Then
            shp.Delete
        End If
    Next shp

    ' Create choice buttons
    Dim i As Long
    Dim btn As Shape
    Dim topCell As Range

    For i = 1 To modConfig.MAX_CHOICES
        Set topCell = ws.Range("B" & (modConfig.CHOICE_START_ROW + i - 1))

        Set btn = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            topCell.Left + 2, _
            topCell.Top + 1, _
            topCell.MergeArea.Width - 4, _
            topCell.MergeArea.Height - 2)

        With btn
            .Name = modConfig.BTN_PREFIX & i
            .TextFrame2.TextRange.Text = ""
            .TextFrame2.TextRange.Font.Size = 11
            .TextFrame2.TextRange.Font.Name = "Georgia"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                RGB(modConfig.C_GOLD_R, modConfig.C_GOLD_G, modConfig.C_GOLD_B)
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .TextFrame2.MarginLeft = 12
            .TextFrame2.MarginRight = 8
            .TextFrame2.MarginTop = 2
            .TextFrame2.MarginBottom = 2
            .TextFrame2.WordWrap = msoTrue

            ' Dark button styling
            .Fill.ForeColor.RGB = _
                RGB(modConfig.C_PANEL_R, modConfig.C_PANEL_G, modConfig.C_PANEL_B)
            .Line.ForeColor.RGB = _
                RGB(modConfig.C_BORDER_R, modConfig.C_BORDER_G, modConfig.C_BORDER_B)
            .Line.Weight = 0.75

            ' Round corners
            .Adjustments.Item(1) = 0.08

            ' Assign click macro
            .OnAction = "ChoiceClicked_" & i

            .Visible = msoFalse  ' Hidden by default
        End With
    Next i

    ' Initialize and start
    StartNewGame

    MsgBox "Damned Moon initialized." & vbCrLf & _
           "Buttons created. Game ready." & vbCrLf & vbCrLf & _
           "Save as .xlsm to preserve macros.", _
           vbInformation, "DAMNED MOON"
End Sub

'===============================================================
' RESET GAME STATE — Dev/debug full reset
'===============================================================
Public Sub ResetGameState()
    If MsgBox("Reset ALL game state? This cannot be undone.", _
              vbYesNo + vbCritical, "RESET") = vbNo Then Exit Sub

    modState.ResetGameState
    modData.InvalidateCaches
    mGameInitialized = False

    MsgBox "Game state fully reset.", vbInformation, "RESET"
End Sub

'===============================================================
' CHOICE CLICK HANDLERS — Entry points for button macros
'===============================================================
Public Sub ChoiceClicked_1()
    EnsureInitialized
    modSceneEngine.ProcessChoice 1
End Sub

Public Sub ChoiceClicked_2()
    EnsureInitialized
    modSceneEngine.ProcessChoice 2
End Sub

Public Sub ChoiceClicked_3()
    EnsureInitialized
    modSceneEngine.ProcessChoice 3
End Sub

Public Sub ChoiceClicked_4()
    EnsureInitialized
    modSceneEngine.ProcessChoice 4
End Sub

Public Sub ChoiceClicked_5()
    EnsureInitialized
    modSceneEngine.ProcessChoice 5
End Sub

'===============================================================
' INTERNAL HELPERS
'===============================================================

' Lazy init — ensures caches are built before any game action
Private Sub EnsureInitialized()
    If Not mGameInitialized Then
        InitializeGame
    End If
End Sub

' Validate that critical sheets exist
Private Sub ValidateCriticalSheets()
    Dim criticalSheets As Variant
    criticalSheets = Array( _
        modConfig.SH_GAME, _
        modConfig.SH_SCENES, _
        modConfig.SH_FLAGS, _
        modConfig.SH_STATS)

    Dim i As Long
    For i = LBound(criticalSheets) To UBound(criticalSheets)
        Dim ws As Worksheet
        Set ws = modConfig.GetSheet(CStr(criticalSheets(i)))
        If ws Is Nothing Then
            modUtils.ErrorLog "modBootstrap.ValidateCriticalSheets", _
                "Missing critical sheet: " & CStr(criticalSheets(i))
        Else
            modUtils.DebugLog "  verified sheet: " & CStr(criticalSheets(i))
        End If
    Next i
End Sub

'===============================================================
' DEV / TEST TOOLS
'===============================================================

' Jump to any scene (dev shortcut)
Public Sub DevJumpToScene(sceneID As String)
    EnsureInitialized
    If Not modData.SceneExists(sceneID) Then
        MsgBox "Scene '" & sceneID & "' not found.", vbExclamation, "DEV"
        Exit Sub
    End If
    modSceneEngine.LoadScene sceneID
    modUtils.DebugLog "modBootstrap.DevJumpToScene: " & sceneID
End Sub

' Print current game state to Immediate window
Public Sub DevPrintState()
    Debug.Print "=== GAME STATE ==="
    Debug.Print "Scene:    " & modState.GetCurrentScene()
    Debug.Print "Location: " & modState.GetCurrentLocation()
    Debug.Print "Day:      " & modState.GetCurrentDay()
    Debug.Print "Time:     " & modState.GetTimeOfDay()
    Debug.Print "Moon:     " & modState.GetMoonPhase()
    Debug.Print "Health:   " & modState.GetStat(modConfig.STAT_HEALTH)
    Debug.Print "Humanity: " & modState.GetStat(modConfig.STAT_HUMANITY)
    Debug.Print "Rage:     " & modState.GetStat(modConfig.STAT_RAGE)
    Debug.Print "Hunger:   " & modState.GetStat(modConfig.STAT_HUNGER)
    Debug.Print "Composure:" & modState.GetStat(modConfig.STAT_COMPOSURE)
    Debug.Print "Instinct: " & modState.GetStat(modConfig.STAT_INSTINCT)
    Debug.Print "Control:  " & modState.GetControl()
    Debug.Print "Danger:   " & modState.GetDangerLevel()
    Debug.Print "Weapon:   " & modState.GetEquippedWeapon()
    Debug.Print "Night:    " & modState.IsNight()
    Debug.Print "FullMoon: " & modState.IsFullMoon()
    Debug.Print "Caches:   " & modData.AreCachesBuilt()
    Debug.Print "================="
End Sub

' Rebuild all caches (dev shortcut)
Public Sub DevRebuildCaches()
    modData.InvalidateCaches
    modData.BuildCaches
    MsgBox "Caches rebuilt.", vbInformation, "DEV"
End Sub

' Set a stat from the Immediate window: DevSetStat "RAGE", 80
Public Sub DevSetStat(statName As String, val As Variant)
    EnsureInitialized
    modState.SetStat statName, val
    Debug.Print "Set " & statName & " = " & CStr(val)
End Sub

' Set a flag from the Immediate window: DevSetFlag "MARIE_MET", True
Public Sub DevSetFlag(flagName As String, val As Boolean)
    EnsureInitialized
    modState.SetFlag flagName, val
    Debug.Print "Set " & flagName & " = " & CStr(val)
End Sub
