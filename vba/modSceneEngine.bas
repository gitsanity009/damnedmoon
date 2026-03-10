Attribute VB_Name = "modSceneEngine"
'===============================================================
' modSceneEngine — Scene Loading & Choice Processing
' Damned Moon VBA RPG Engine — Phase 2
'===============================================================
' The core game loop: load a scene, display its narrative and
' choices, process the player's selection, apply effects, and
' transition to the next scene.
'
' Delegates to:
'   modRequirements — choice availability checks
'   modEffects      — stat/flag/item/quest mutations
'   modUI           — rendering narrative, choices, HUD panels
'===============================================================

Option Explicit

'===============================================================
' LOAD SCENE — Display a scene and its choices
'===============================================================
Public Sub LoadScene(sceneID As String)
    Application.ScreenUpdating = False

    Dim wsScenes As Worksheet
    Set wsScenes = modConfig.GetSheet(modConfig.SH_SCENES)
    If wsScenes Is Nothing Then
        modUtils.ErrorLog "modSceneEngine.LoadScene", "Scenes sheet not found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Find scene row via cache
    Dim sceneRow As Long
    sceneRow = modData.GetSceneRow(sceneID)
    If sceneRow = 0 Then
        modUI.ShowNarrative "[ERROR: Scene " & sceneID & " not found]"
        modUtils.ErrorLog "modSceneEngine.LoadScene", "Scene '" & sceneID & "' not in cache"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Update current scene in state
    modState.SetCurrentScene sceneID

    ' Process OnEnter effects (column AA)
    Dim enterEffects As String
    enterEffects = modUtils.SafeStr(wsScenes.Cells(sceneRow, modConfig.SCN_COL_ONENTER).Value)
    If Len(enterEffects) > 0 Then
        Dim jumpFromEnter As String
        jumpFromEnter = modEffects.ProcessEffects(enterEffects)
        ' If OnEnter triggers a scene jump, redirect
        If Len(jumpFromEnter) > 0 Then
            Application.ScreenUpdating = True
            LoadScene jumpFromEnter
            Exit Sub
        End If
    End If

    ' Update location from scene data
    Dim loc As String
    loc = modUtils.SafeStr(wsScenes.Cells(sceneRow, modConfig.SCN_COL_LOCATION).Value)
    If Len(loc) > 0 Then
        modState.SetCurrentLocation loc
    End If

    ' Update time from scene data
    Dim timeSlot As String
    timeSlot = modUtils.SafeStr(wsScenes.Cells(sceneRow, modConfig.SCN_COL_TIME).Value)
    If Len(timeSlot) > 0 Then
        modState.SetTimeOfDay timeSlot
    End If

    ' Update day from scene data
    Dim dayRange As String
    dayRange = modUtils.SafeStr(wsScenes.Cells(sceneRow, modConfig.SCN_COL_DAY).Value)
    If Len(dayRange) > 0 And IsNumeric(Left(dayRange, 1)) Then
        Dim dayNum As Long
        Dim plusPos As Long
        plusPos = InStr(dayRange, "+")
        If plusPos > 0 Then
            dayNum = modUtils.SafeLng(Left(dayRange, plusPos - 1), 0)
        Else
            dayNum = modUtils.SafeLng(dayRange, 0)
        End If
        If dayNum > modState.GetCurrentDay() Then
            modState.SetCurrentDay dayNum
        End If
    End If

    ' Load narrative text
    Dim narrative As String
    narrative = modUtils.SafeStr(wsScenes.Cells(sceneRow, modConfig.SCN_COL_NARRATIVE).Value)
    modUI.ShowNarrative narrative

    ' Load and render choices
    LoadChoices sceneRow

    ' Update all HUD panels
    modUI.UpdateStatsPanel
    modUI.UpdateQuestPanel
    modUI.UpdateInventoryPanel
    modUI.UpdateDayTimePanel
    modUI.UpdateMapHighlight loc

    ' Check for quest progression
    CheckQuestProgress sceneID

    ' Check for scene-triggered combat (column AC)
    Dim combatEnemy As String
    combatEnemy = modUtils.SafeStr(wsScenes.Cells(sceneRow, modConfig.SCN_COL_COMBAT).Value)
    If Len(combatEnemy) > 0 Then
        Dim combatResult As String
        combatResult = modCombat.QuickCombat(combatEnemy)
        modJournal.AddCombatEntry "Fought " & modCombat.GetEnemyDisplayName(combatEnemy) & " — " & combatResult
        If combatResult = modCombat.RESULT_DEFEAT Then
            Application.ScreenUpdating = True
            LoadScene "SCN_DEFEAT"
            Exit Sub
        End If
    End If

    ' Auto-save on scene load
    modSave.AutoSave

    modUtils.DebugLog "modSceneEngine.LoadScene: loaded " & sceneID
    Application.ScreenUpdating = True
End Sub

'===============================================================
' PROCESS CHOICE — Handle a player clicking choice N
'===============================================================
Public Sub ProcessChoice(choiceNum As Long)
    Dim wsScenes As Worksheet
    Set wsScenes = modConfig.GetSheet(modConfig.SH_SCENES)
    If wsScenes Is Nothing Then Exit Sub

    ' Get current scene
    Dim currentScene As String
    currentScene = modState.GetCurrentScene()
    If Len(currentScene) = 0 Then Exit Sub

    Dim sceneRow As Long
    sceneRow = modData.GetSceneRow(currentScene)
    If sceneRow = 0 Then Exit Sub

    ' Calculate choice columns: C1=G/H/I/J, C2=K/L/M/N, etc.
    Dim baseCol As Long
    baseCol = modConfig.CHOICE_BASE_COL + (choiceNum - 1) * modConfig.CHOICE_COL_SPAN

    ' Verify choice exists
    Dim choiceText As String
    choiceText = modUtils.SafeStr(wsScenes.Cells(sceneRow, baseCol).Value)
    If Len(choiceText) = 0 Then Exit Sub

    ' Check requirements (column baseCol + 2)
    Dim reqStr As String
    reqStr = modUtils.SafeStr(wsScenes.Cells(sceneRow, baseCol + 2).Value)
    If Len(reqStr) > 0 Then
        If Not modRequirements.CheckRequirements(reqStr) Then
            ' Requirement not met — flash locked
            modUI.FlashButton choiceNum, False
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False

    ' Push undo snapshot before applying choice
    modSave.PushSnapshot

    ' Process choice effects (column baseCol + 3)
    Dim effectStr As String
    effectStr = modUtils.SafeStr(wsScenes.Cells(sceneRow, baseCol + 3).Value)
    Dim jumpFromChoice As String
    If Len(effectStr) > 0 Then
        jumpFromChoice = modEffects.ProcessEffects(effectStr)
    End If

    ' Process OnExit effects of current scene (column AB)
    Dim exitEffects As String
    exitEffects = modUtils.SafeStr(wsScenes.Cells(sceneRow, modConfig.SCN_COL_ONEXIT).Value)
    If Len(exitEffects) > 0 Then
        Dim jumpFromExit As String
        jumpFromExit = modEffects.ProcessEffects(exitEffects)
        If Len(jumpFromExit) > 0 And Len(jumpFromChoice) = 0 Then
            jumpFromChoice = jumpFromExit
        End If
    End If

    Application.ScreenUpdating = True

    ' Determine target scene
    Dim targetScene As String
    If Len(jumpFromChoice) > 0 Then
        ' Effect-triggered jump overrides normal target
        targetScene = jumpFromChoice
    Else
        ' Normal target from scene table (column baseCol + 1)
        targetScene = modUtils.SafeStr(wsScenes.Cells(sceneRow, baseCol + 1).Value)
    End If

    ' Load next scene
    If Len(targetScene) > 0 Then
        LoadScene targetScene
    End If
End Sub

'===============================================================
' LOAD CHOICES — Read choices from scene row, render buttons
'===============================================================
Private Sub LoadChoices(sceneRow As Long)
    Dim wsScenes As Worksheet
    Set wsScenes = modConfig.GetSheet(modConfig.SH_SCENES)
    If wsScenes Is Nothing Then Exit Sub

    Dim choiceCount As Long
    choiceCount = 0

    Dim i As Long
    For i = 1 To modConfig.MAX_CHOICES
        Dim baseCol As Long
        baseCol = modConfig.CHOICE_BASE_COL + (i - 1) * modConfig.CHOICE_COL_SPAN

        Dim choiceText As String
        choiceText = modUtils.SafeStr(wsScenes.Cells(sceneRow, baseCol).Value)

        If Len(choiceText) > 0 Then
            ' Check requirements for visual state
            Dim reqStr As String
            reqStr = modUtils.SafeStr(wsScenes.Cells(sceneRow, baseCol + 2).Value)

            Dim isAvailable As Boolean
            If Len(reqStr) > 0 Then
                isAvailable = modRequirements.CheckRequirements(reqStr)
            Else
                isAvailable = True
            End If

            ' Build display text
            Dim displayText As String
            displayText = CStr(i) & ".  " & choiceText

            ' If locked, append reason
            If Not isAvailable Then
                Dim reason As String
                reason = modRequirements.GetFailReason(reqStr)
                If Len(reason) > 0 Then
                    displayText = displayText & "  " & reason
                End If
            End If

            ' Show the button
            modUI.ShowChoiceButton i, displayText, isAvailable
            choiceCount = choiceCount + 1
        Else
            ' No choice — hide button
            modUI.HideChoiceButton i
        End If
    Next i

    ' Hide any remaining buttons
    Dim j As Long
    For j = choiceCount + 1 To modConfig.MAX_CHOICES
        ' Already handled above if choiceText was empty
    Next j

    ' Store choice count
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If Not ws Is Nothing Then
        ws.Range(modConfig.CHOICE_COUNT_CELL).Value = choiceCount
    End If
End Sub

'===============================================================
' QUEST PROGRESSION — Check if scene triggers quest advancement
'===============================================================
Private Sub CheckQuestProgress(sceneID As String)
    Dim wsQ As Worksheet, wsQS As Worksheet
    Set wsQ = modConfig.GetSheet(modConfig.SH_QUESTS)
    Set wsQS = modConfig.GetSheet(modConfig.SH_QUESTSTAGES)
    If wsQ Is Nothing Or wsQS Is Nothing Then Exit Sub

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsQS, 1)
        Dim questID As String
        questID = modUtils.SafeStr(wsQS.Cells(r, 1).Value)
        Dim stageIdx As Long
        stageIdx = modUtils.SafeLng(wsQS.Cells(r, 2).Value, -1)
        Dim trigger As String
        trigger = modUtils.SafeStr(wsQS.Cells(r, 5).Value)
        Dim triggerType As String
        triggerType = modUtils.SafeStr(wsQS.Cells(r, 6).Value)

        Dim shouldAdvance As Boolean
        shouldAdvance = False

        If triggerType = "SCENE_COMPLETE" And trigger = sceneID Then
            shouldAdvance = True
        ElseIf triggerType = "FLAG_SET" Then
            shouldAdvance = modState.GetFlag(trigger)
        End If

        If shouldAdvance Then
            Dim qRow As Long
            qRow = modData.GetQuestRow(questID)
            If qRow > 0 Then
                Dim currentStage As Long
                currentStage = modUtils.SafeLng(wsQ.Cells(qRow, 6).Value, -1)

                If stageIdx = currentStage + 1 Or (currentStage = -1 And stageIdx = 0) Then
                    wsQ.Cells(qRow, 6).Value = stageIdx
                    wsQ.Cells(qRow, 4).Value = "ACTIVE"
                    wsQ.Cells(qRow, 5).Value = modUtils.SafeStr(wsQS.Cells(r, 4).Value)

                    ' Award XP
                    Dim xp As Long
                    xp = modUtils.SafeLng(wsQS.Cells(r, 8).Value, 0)
                    If xp > 0 Then
                        modState.AddStat modConfig.STAT_XP, xp
                    End If

                    modUtils.DebugLog "modSceneEngine.CheckQuestProgress: " & questID & " -> stage " & stageIdx
                End If
            End If
        End If
    Next r
End Sub
