Attribute VB_Name = "DamnedMoon_Engine"
'===============================================================
' DAMNED MOON — VBA GAME ENGINE v2.0
' Excel/VBA Choose-Your-Own-Adventure RPG
' Wroughtwood, 1960
'===============================================================
' SETUP: Import this module via Developer > Visual Basic > File > Import
' Then run SetupGame() once to create buttons and initialize.
' Save as .xlsm (macro-enabled workbook).
'===============================================================

Option Explicit

' ── SHEET REFERENCES ──
Private Const SH_GAME As String = "Game"
Private Const SH_SCENES As String = "tbl_Scenes"
Private Const SH_FLAGS As String = "tbl_Flags"
Private Const SH_STATS As String = "Stats"
Private Const SH_ITEMS As String = "tbl_ItemDB"
Private Const SH_INV As String = "tbl_Inventory"
Private Const SH_QUESTS As String = "tbl_Quests"
Private Const SH_QUESTSTAGES As String = "tbl_QuestStages"
Private Const SH_ENEMIES As String = "tbl_Enemies"
Private Const SH_MOON As String = "tbl_MoonPhases"
Private Const SH_JOBS As String = "tbl_Jobs"
Private Const SH_COMBAT As String = "tbl_CombatLog"
Private Const SH_SAVES As String = "SaveSlots"
Private Const SH_CONFIG As String = "Config"

' ── GAME SHEET LAYOUT CONSTANTS ──
Private Const NARRATIVE_CELL As String = "B6"
Private Const SCENE_ID_CELL As String = "E40"
Private Const CHOICE_COUNT_CELL As String = "E41"
Private Const LOCATION_CELL As String = "E42"
Private Const CHOICE_START_ROW As Long = 25
Private Const CHOICE_END_ROW As Long = 29
Private Const DAY_CELL As String = "E2"
Private Const TIME_CELL As String = "E3"
Private Const MOON_CELL As String = "H2"
Private Const MAP_LOCATION_CELL As String = "L3"

' ── MAP HIGHLIGHT RANGES (building cells to highlight) ──
Private Type MapLocation
    Name As String
    LocationCode As String
    CellRange As String
    Row As Long
    Col As Long
End Type

' ── BUTTON NAMES ──
Private Const BTN_PREFIX As String = "btnChoice"

'===============================================================
' SETUP — Run this ONCE to create buttons and initialize
'===============================================================
Public Sub SetupGame()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    
    ' Delete old buttons if they exist
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, Len(BTN_PREFIX)) = BTN_PREFIX Then
            shp.Delete
        End If
    Next shp
    
    ' Create 5 choice buttons
    Dim i As Long
    Dim btn As Shape
    Dim topCell As Range
    
    For i = 1 To 5
        Set topCell = ws.Range("B" & (CHOICE_START_ROW + i - 1))
        
        Set btn = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            topCell.Left + 2, _
            topCell.Top + 1, _
            topCell.MergeArea.Width - 4, _
            topCell.MergeArea.Height - 2)
        
        With btn
            .Name = BTN_PREFIX & i
            .TextFrame2.TextRange.Text = ""
            .TextFrame2.TextRange.Font.Size = 11
            .TextFrame2.TextRange.Font.Name = "Georgia"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(201, 162, 39) ' Gold
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .TextFrame2.MarginLeft = 12
            .TextFrame2.MarginRight = 8
            .TextFrame2.MarginTop = 2
            .TextFrame2.MarginBottom = 2
            .TextFrame2.WordWrap = msoTrue
            
            ' Dark button styling
            .Fill.ForeColor.RGB = RGB(34, 26, 18)  ' C_PANEL2
            .Line.ForeColor.RGB = RGB(58, 46, 34)  ' C_BORDER
            .Line.Weight = 0.75
            
            ' Round corners
            .Adjustments.Item(1) = 0.08
            
            ' Assign click macro
            .OnAction = "ChoiceClicked_" & i
            
            .Visible = msoFalse  ' Hidden by default
        End With
    Next i
    
    ' Initialize game state
    LoadScene "SCN_PROLOGUE"
    
    MsgBox "Damned Moon initialized." & vbCrLf & _
           "Buttons created. Game ready." & vbCrLf & vbCrLf & _
           "Save as .xlsm to preserve macros.", _
           vbInformation, "DAMNED MOON"
End Sub

'===============================================================
' CHOICE CLICK HANDLERS
'===============================================================
Public Sub ChoiceClicked_1()
    ProcessChoice 1
End Sub
Public Sub ChoiceClicked_2()
    ProcessChoice 2
End Sub
Public Sub ChoiceClicked_3()
    ProcessChoice 3
End Sub
Public Sub ChoiceClicked_4()
    ProcessChoice 4
End Sub
Public Sub ChoiceClicked_5()
    ProcessChoice 5
End Sub

'===============================================================
' CORE GAME ENGINE
'===============================================================
Public Sub ProcessChoice(choiceNum As Long)
    Dim ws As Worksheet, wsScenes As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    Set wsScenes = ThisWorkbook.Sheets(SH_SCENES)
    
    Dim currentScene As String
    currentScene = ws.Range(SCENE_ID_CELL).Value
    
    ' Find scene row
    Dim sceneRow As Long
    sceneRow = FindSceneRow(currentScene)
    If sceneRow = 0 Then Exit Sub
    
    ' Choice columns: C1=G/H/I/J, C2=K/L/M/N, C3=O/P/Q/R, C4=S/T/U/V, C5=W/X/Y/Z
    Dim baseCol As Long
    baseCol = 7 + (choiceNum - 1) * 4  ' G=7, K=11, O=15, S=19, W=23
    
    Dim choiceText As String
    choiceText = CStr(wsScenes.Cells(sceneRow, baseCol).Value & "")
    If choiceText = "" Then Exit Sub
    
    ' Get target scene
    Dim targetScene As String
    targetScene = CStr(wsScenes.Cells(sceneRow, baseCol + 1).Value & "")
    
    ' Check requirements (column baseCol + 2)
    Dim reqStr As String
    reqStr = CStr(wsScenes.Cells(sceneRow, baseCol + 2).Value & "")
    If reqStr <> "" Then
        If Not CheckRequirement(reqStr) Then
            ' Requirement not met — flash the button red briefly
            FlashButton choiceNum, False
            Exit Sub
        End If
    End If
    
    ' Process choice effects (column baseCol + 3)
    Dim effectStr As String
    effectStr = CStr(wsScenes.Cells(sceneRow, baseCol + 3).Value & "")
    If effectStr <> "" Then
        ProcessEffects effectStr
    End If
    
    ' Process OnExit effects of current scene (column AA = 27)
    Dim exitEffects As String
    exitEffects = CStr(wsScenes.Cells(sceneRow, 28).Value & "")
    If exitEffects <> "" Then
        ProcessEffects exitEffects
    End If
    
    ' Load next scene
    If targetScene <> "" Then
        LoadScene targetScene
    End If
End Sub

Public Sub LoadScene(sceneID As String)
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet, wsScenes As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    Set wsScenes = ThisWorkbook.Sheets(SH_SCENES)
    
    ' Find scene
    Dim sceneRow As Long
    sceneRow = FindSceneRow(sceneID)
    If sceneRow = 0 Then
        ws.Range(NARRATIVE_CELL).Value = "[ERROR: Scene " & sceneID & " not found]"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Update current scene ID
    ws.Range(SCENE_ID_CELL).Value = sceneID
    
    ' Process OnEnter effects (column AA = 27)
    Dim enterEffects As String
    enterEffects = CStr(wsScenes.Cells(sceneRow, 27).Value & "")
    If enterEffects <> "" Then
        ProcessEffects enterEffects
    End If
    
    ' Load narrative text
    Dim narrative As String
    narrative = CStr(wsScenes.Cells(sceneRow, 6).Value & "")  ' Column F
    ws.Range(NARRATIVE_CELL).Value = narrative
    
    ' Update location
    Dim loc As String
    loc = CStr(wsScenes.Cells(sceneRow, 3).Value & "")  ' Column C
    ws.Range(LOCATION_CELL).Value = loc
    
    ' Update time
    Dim timeSlot As String
    timeSlot = CStr(wsScenes.Cells(sceneRow, 5).Value & "")  ' Column E
    If timeSlot <> "" Then
        UpdateTimeStat timeSlot
    End If
    
    ' Update day
    Dim dayRange As String
    dayRange = CStr(wsScenes.Cells(sceneRow, 4).Value & "")  ' Column D
    If dayRange <> "" And IsNumeric(Left(dayRange, 1)) Then
        Dim dayNum As Long
        dayNum = CLng(Left(dayRange, InStr(dayRange & "+", "+") - 1))
        If dayNum > GetStat("DAY_COUNTER") Then
            SetStat "DAY_COUNTER", dayNum
        End If
    End If
    
    ' Load choices into buttons
    LoadChoices sceneRow
    
    ' Update all UI panels
    UpdateStatsPanel
    UpdateQuestPanel
    UpdateInventoryPanel
    UpdateDayTimePanel
    UpdateMapHighlight loc
    
    ' Check for combat encounter (column AC = 29)
    Dim combatEnemy As String
    combatEnemy = CStr(wsScenes.Cells(sceneRow, 29).Value & "")
    If combatEnemy <> "" Then
        ' Combat scenes show enemy info in narrative
        ' Full combat system can be extended here
    End If
    
    ' Advance quest stages
    CheckQuestProgress sceneID
    
    Application.ScreenUpdating = True
End Sub

Private Sub LoadChoices(sceneRow As Long)
    Dim ws As Worksheet, wsScenes As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    Set wsScenes = ThisWorkbook.Sheets(SH_SCENES)
    
    Dim choiceCount As Long
    choiceCount = 0
    
    Dim i As Long
    For i = 1 To 5
        Dim baseCol As Long
        baseCol = 7 + (i - 1) * 4
        
        Dim choiceText As String
        choiceText = CStr(wsScenes.Cells(sceneRow, baseCol).Value & "")
        
        Dim btn As Shape
        On Error Resume Next
        Set btn = ws.Shapes(BTN_PREFIX & i)
        On Error GoTo 0
        
        If btn Is Nothing Then GoTo NextChoice
        
        If choiceText <> "" Then
            btn.Visible = msoTrue
            btn.TextFrame2.TextRange.Text = CStr(i) & ".  " & choiceText
            
            ' Check if requirement exists and color accordingly
            Dim reqStr As String
            reqStr = CStr(wsScenes.Cells(sceneRow, baseCol + 2).Value & "")
            
            If reqStr <> "" Then
                If CheckRequirement(reqStr) Then
                    ' Requirement met — gold text
                    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(201, 162, 39)
                    btn.Fill.ForeColor.RGB = RGB(34, 26, 18)
                Else
                    ' Requirement NOT met — dim/locked
                    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(80, 70, 60)
                    btn.Fill.ForeColor.RGB = RGB(20, 16, 12)
                End If
            Else
                ' No requirement — standard gold
                btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(201, 162, 39)
                btn.Fill.ForeColor.RGB = RGB(34, 26, 18)
            End If
            
            choiceCount = choiceCount + 1
        Else
            btn.Visible = msoFalse
        End If
        
NextChoice:
        Set btn = Nothing
    Next i
    
    ws.Range(CHOICE_COUNT_CELL).Value = choiceCount
End Sub

'===============================================================
' EFFECT PROCESSING
'===============================================================
Private Sub ProcessEffects(effectStr As String)
    ' Effects are pipe-delimited: STAT:HEALTH+5|FLAG_SET:MY_FLAG|ITEM_ADD:ITM_KNIFE
    Dim effects() As String
    effects = Split(effectStr, "|")
    
    Dim i As Long
    For i = LBound(effects) To UBound(effects)
        Dim ef As String
        ef = Trim(effects(i))
        If ef = "" Then GoTo NextEffect
        
        If Left(ef, 5) = "STAT:" Then
            ProcessStatEffect Mid(ef, 6)
        ElseIf Left(ef, 9) = "FLAG_SET:" Then
            SetFlag Mid(ef, 10), True
        ElseIf Left(ef, 11) = "FLAG_CLEAR:" Then
            SetFlag Mid(ef, 12), False
        ElseIf Left(ef, 9) = "ITEM_ADD:" Then
            AddItem Mid(ef, 10)
        ElseIf Left(ef, 12) = "ITEM_REMOVE:" Then
            RemoveItem Mid(ef, 13)
        End If
NextEffect:
    Next i
End Sub

Private Sub ProcessStatEffect(statStr As String)
    ' Format: HEALTH+5, RAGE-10, HUNGER=50
    Dim statName As String
    Dim opPos As Long
    Dim op As String
    Dim val As Long
    
    ' Find operator
    Dim j As Long
    For j = 1 To Len(statStr)
        Dim ch As String
        ch = Mid(statStr, j, 1)
        If ch = "+" Or ch = "-" Or ch = "=" Then
            statName = Left(statStr, j - 1)
            op = ch
            val = CLng(Mid(statStr, j + 1))
            Exit For
        End If
    Next j
    
    If statName = "" Then Exit Sub
    
    Dim current As Long
    current = GetStat(statName)
    
    Select Case op
        Case "+": current = current + val
        Case "-": current = current - val
        Case "=": current = val
    End Select
    
    ' Clamp stats
    Select Case statName
        Case "HEALTH": If current > 100 Then current = 100: If current < 0 Then current = 0
        Case "HUMANITY": If current > 100 Then current = 100: If current < 0 Then current = 0
        Case "RAGE": If current > 100 Then current = 100: If current < 0 Then current = 0
        Case "HUNGER": If current > 100 Then current = 100: If current < 0 Then current = 0
    End Select
    
    SetStat statName, current
    
    ' Check RAGE 100 blackout trigger
    If statName = "RAGE" And current >= 100 Then
        TriggerBlackout
    End If
End Sub

'===============================================================
' STAT SYSTEM
'===============================================================
Public Function GetStat(statName As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_STATS)
    
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = statName Then
            Dim v As Variant
            v = ws.Cells(r, 3).Value
            If IsNumeric(v) Then
                GetStat = CLng(v)
            Else
                GetStat = 0
            End If
            Exit Function
        End If
    Next r
    GetStat = 0
End Function

' Use this for text-valued stats like TIME_OF_DAY and MOON_PHASE
Public Function GetStatText(statName As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_STATS)
    
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = statName Then
            GetStatText = CStr(ws.Cells(r, 3).Value & "")
            Exit Function
        End If
    Next r
    GetStatText = ""
End Function

Public Sub SetStat(statName As String, val As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_STATS)
    
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = statName Then
            ws.Cells(r, 3).Value = val
            Exit Sub
        End If
    Next r
End Sub

Private Sub UpdateTimeStat(timeSlot As String)
    SetStat "TIME_OF_DAY", timeSlot
End Sub

'===============================================================
' FLAG SYSTEM
'===============================================================
Public Function GetFlag(flagName As String) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_FLAGS)
    
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = flagName Then
            Dim v As Variant
            v = ws.Cells(r, 2).Value
            If IsEmpty(v) Or CStr(v & "") = "" Then
                GetFlag = False
            ElseIf IsNumeric(v) Then
                GetFlag = (CLng(v) <> 0)
            Else
                GetFlag = CBool(v)
            End If
            Exit Function
        End If
    Next r
    GetFlag = False
End Function

Public Sub SetFlag(flagName As String, val As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_FLAGS)
    
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = flagName Then
            ws.Cells(r, 2).Value = val
            Exit Sub
        End If
    Next r
End Sub

Private Function CheckRequirement(reqStr As String) As Boolean
    ' Formats: FLAG:MY_FLAG, STAT:RAGE>50, STAT:HUMANITY<50
    If Left(reqStr, 5) = "FLAG:" Then
        CheckRequirement = GetFlag(Mid(reqStr, 6))
    ElseIf Left(reqStr, 5) = "STAT:" Then
        Dim expr As String
        expr = Mid(reqStr, 6)
        ' Parse: RAGE>50
        Dim opPos As Long
        Dim k As Long
        For k = 1 To Len(expr)
            Dim c As String
            c = Mid(expr, k, 1)
            If c = ">" Or c = "<" Or c = "=" Then
                Dim sName As String
                Dim sOp As String
                Dim sVal As Long
                sName = Left(expr, k - 1)
                sOp = c
                sVal = CLng(Mid(expr, k + 1))
                
                Dim current As Long
                current = GetStat(sName)
                
                Select Case sOp
                    Case ">": CheckRequirement = (current > sVal)
                    Case "<": CheckRequirement = (current < sVal)
                    Case "=": CheckRequirement = (current = sVal)
                End Select
                Exit Function
            End If
        Next k
        CheckRequirement = True
    Else
        CheckRequirement = True
    End If
End Function

'===============================================================
' INVENTORY SYSTEM
'===============================================================
Public Sub AddItem(itemID As String)
    Dim wsInv As Worksheet, wsItems As Worksheet
    Set wsInv = ThisWorkbook.Sheets(SH_INV)
    Set wsItems = ThisWorkbook.Sheets(SH_ITEMS)
    
    ' Find item in DB
    Dim itemName As String
    Dim r As Long
    For r = 2 To wsItems.Cells(wsItems.Rows.Count, 1).End(xlUp).Row
        If CStr(wsItems.Cells(r, 1).Value) = itemID Then
            itemName = CStr(wsItems.Cells(r, 2).Value)
            Exit For
        End If
    Next r
    
    If itemName = "" Then Exit Sub
    
    ' Check if already in inventory (stack)
    Dim ir As Long
    For ir = 2 To wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
        If CStr(wsInv.Cells(ir, 2).Value) = itemID Then
            wsInv.Cells(ir, 4).Value = CLng(wsInv.Cells(ir, 4).Value) + 1
            Exit Sub
        End If
    Next ir
    
    ' Find empty slot
    For ir = 2 To wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
        Dim slotQty As Variant
        slotQty = wsInv.Cells(ir, 4).Value
        If CStr(wsInv.Cells(ir, 2).Value & "") = "" And (IsEmpty(slotQty) Or slotQty = 0) Then
            wsInv.Cells(ir, 2).Value = itemID
            wsInv.Cells(ir, 3).Value = itemName
            wsInv.Cells(ir, 4).Value = 1
            Exit Sub
        End If
    Next ir
End Sub

Public Sub RemoveItem(itemID As String)
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets(SH_INV)
    
    Dim r As Long
    For r = 2 To wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
        If CStr(wsInv.Cells(r, 2).Value) = itemID Then
            Dim qty As Long
            qty = CLng(wsInv.Cells(r, 4).Value) - 1
            If qty <= 0 Then
                wsInv.Cells(r, 2).Value = ""
                wsInv.Cells(r, 3).Value = ""
                wsInv.Cells(r, 4).Value = 0
            Else
                wsInv.Cells(r, 4).Value = qty
            End If
            Exit Sub
        End If
    Next r
End Sub

'===============================================================
' UI UPDATE FUNCTIONS
'===============================================================
Private Sub UpdateStatsPanel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    
    ' Stat display rows 7-12 in columns E-F
    Dim stats As Variant
    stats = Array("HEALTH", "HUMANITY", "RAGE", "HUNGER", "COMPOSURE", "INSTINCT")
    
    Dim i As Long
    For i = 0 To 5
        Dim val As Long
        val = GetStat(CStr(stats(i)))
        ws.Cells(7 + i, 6).Value = val  ' Column F
    Next i
    
    ' HP display
    Dim hp As Long
    hp = GetStat("HEALTH")
    ws.Range("E15").Value = "  HP: " & hp & " / 100"
    
    ' Day display
    ws.Range("E2").Value = "DAY " & GetStat("DAY_COUNTER")
    ws.Range("E3").Value = CStr(GetStat("TIME_OF_DAY") & "")
    
    ' Moon phase
    UpdateMoonDisplay
End Sub

Private Sub UpdateMoonDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    Dim wsMoon As Worksheet
    Set wsMoon = ThisWorkbook.Sheets(SH_MOON)
    
    Dim day As Long
    day = GetStat("DAY_COUNTER")
    
    Dim r As Long
    For r = 2 To wsMoon.Cells(wsMoon.Rows.Count, 1).End(xlUp).Row
        Dim dayRange As String
        dayRange = CStr(wsMoon.Cells(r, 2).Value)
        
        ' Parse day range (e.g., "1-4", "30")
        If InStr(dayRange, "-") > 0 Then
            Dim parts() As String
            parts = Split(dayRange, "-")
            If day >= CLng(parts(0)) And day <= CLng(parts(1)) Then
                Dim phase As String
                phase = CStr(wsMoon.Cells(r, 1).Value)
                Dim moonSym As String
                Select Case phase
                    Case "New Moon": moonSym = ChrW(&H25CB)  ' empty circle
                    Case "Waxing Crescent": moonSym = ChrW(&H25D1)
                    Case "First Quarter": moonSym = ChrW(&H25D1)
                    Case "Waxing Gibbous": moonSym = ChrW(&H25D0)
                    Case "Full Moon", "Full Moon (Night 30)": moonSym = ChrW(&H25CF)
                    Case "Waning Gibbous": moonSym = ChrW(&H25D0)
                    Case "Last Quarter": moonSym = ChrW(&H25D1)
                    Case "Waning Crescent": moonSym = ChrW(&H25D1)
                    Case Else: moonSym = ChrW(&H25CB)
                End Select
                ws.Range(MOON_CELL).Value = moonSym & " " & UCase(phase)
                Exit Sub
            End If
        ElseIf IsNumeric(dayRange) Then
            If day = CLng(dayRange) Then
                ws.Range(MOON_CELL).Value = ChrW(&H25CF) & " FULL MOON"
                Exit Sub
            End If
        End If
    Next r
End Sub

Private Sub UpdateQuestPanel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Sheets(SH_QUESTS)
    
    Dim questText As String
    Dim r As Long
    For r = 2 To wsQ.Cells(wsQ.Rows.Count, 1).End(xlUp).Row
        If CStr(wsQ.Cells(r, 4).Value) = "ACTIVE" Then
            Dim qName As String
            qName = CStr(wsQ.Cells(r, 2).Value)
            Dim qDesc As String
            qDesc = CStr(wsQ.Cells(r, 5).Value)
            questText = questText & ChrW(&H25C6) & " " & qName & vbLf & _
                        "   " & qDesc & vbLf & vbLf
        End If
    Next r
    
    If questText = "" Then questText = "(No active quests)"
    ws.Range("E18").Value = questText
End Sub

Private Sub UpdateInventoryPanel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets(SH_INV)
    
    ' Update weapon display
    Dim weaponName As String
    weaponName = "(none)"
    
    Dim r As Long
    For r = 2 To wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
        Dim eqVal As Variant
        eqVal = wsInv.Cells(r, 5).Value
        If Not IsEmpty(eqVal) And CStr(eqVal & "") <> "" Then
            If CBool(eqVal) Then
                weaponName = CStr(wsInv.Cells(r, 3).Value)
                Exit For
            End If
        End If
    Next r
    ws.Range("H6").Value = ChrW(&H2694) & " Weapon: " & weaponName
    
    ' Update inventory slots
    Dim slotIdx As Long
    slotIdx = 0
    For r = 2 To wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
        Dim iName As String
        iName = CStr(wsInv.Cells(r, 3).Value & "")
        If iName <> "" Then
            If slotIdx < 5 Then
                ws.Cells(7 + slotIdx, 8).Value = "  " & iName  ' Column H
                Dim qty As Long
                qty = CLng(wsInv.Cells(r, 4).Value)
                If qty > 1 Then
                    ws.Cells(7 + slotIdx, 9).Value = "x" & qty  ' Column I
                Else
                    ws.Cells(7 + slotIdx, 9).Value = ""
                End If
                slotIdx = slotIdx + 1
            End If
        End If
    Next r
    
    ' Clear remaining slots
    Dim s As Long
    For s = slotIdx To 4
        ws.Cells(7 + s, 8).Value = "  [ Empty ]"
        ws.Cells(7 + s, 9).Value = ""
    Next s
End Sub

Private Sub UpdateDayTimePanel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    
    ws.Range("E2").Value = "DAY " & GetStat("DAY_COUNTER")
    ws.Range("E3").Value = GetStatText("TIME_OF_DAY")
End Sub

'===============================================================
' MAP HIGHLIGHTING
'===============================================================
Private Sub UpdateMapHighlight(locationCode As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    
    ' Update location label
    Dim locLabel As String
    Select Case locationCode
        Case "LOS_ANGELES": locLabel = "LOS ANGELES"
        Case "FIELD", "FIELD_EDGE": locLabel = "THE FIELD"
        Case "ROAD": locLabel = "NORTH ROAD"
        Case "MAIN_ROAD": locLabel = "MAIN ROAD"
        Case "INN", "INN_ROOM", "INN_KITCHEN": locLabel = "HATTIE'S INN"
        Case "FEED_STORE": locLabel = "CALHOUN'S FEED STORE"
        Case "SHERIFF_OFFICE": locLabel = "SHERIFF'S OFFICE"
        Case "CHURCH", "CHURCH_OFFICE": locLabel = "WROUGHTWOOD CHURCH"
        Case "SWAMP", "SWAMP_CABIN": locLabel = "MARIE'S SWAMP"
        Case "BARN": locLabel = "THE BARN"
        Case "MARSH_HOUSE": locLabel = "MARSH HOUSE"
        Case "NORTH_ROAD": locLabel = "NORTH ROAD"
        Case "NORTH_WOODS", "DEEP_NORTH_WOODS": locLabel = "NORTH WOODS"
        Case "PACK_CLEARING": locLabel = "PACK CLEARING"
        Case "TREE_LINE": locLabel = "TREE LINE"
        Case "MILL": locLabel = "THE MILL"
        Case Else: locLabel = UCase(locationCode)
    End Select
    
    ws.Range(MAP_LOCATION_CELL).Value = "Current: " & locLabel
    
    ' Reset all building highlights to default
    ResetMapHighlights
    
    ' Highlight current location building
    Dim hlRange As String
    Select Case locationCode
        Case "FIELD", "FIELD_EDGE": hlRange = "O9:U9"
        Case "ROAD", "NORTH_ROAD": hlRange = "O11:U11"
        Case "INN", "INN_ROOM", "INN_KITCHEN": hlRange = "L16:P17"
        Case "FEED_STORE": hlRange = "S13:V14"
        Case "SHERIFF_OFFICE": hlRange = "S16:V17"
        Case "CHURCH", "CHURCH_OFFICE": hlRange = "L13:N14"
        Case "SWAMP", "SWAMP_CABIN": hlRange = "M27:P28"
        Case "BARN": hlRange = "W19:Y20"
        Case "MARSH_HOUSE": hlRange = "L22:O23"
        Case "NORTH_WOODS", "DEEP_NORTH_WOODS": hlRange = "N4:W4"
        Case "PACK_CLEARING": hlRange = "L5:N6"
        Case "TREE_LINE": hlRange = "O8:V8"
        Case "MAIN_ROAD": hlRange = "Q13:R14"
        Case Else: hlRange = ""
    End Select
    
    If hlRange <> "" Then
        Dim rng As Range
        Set rng = ws.Range(hlRange)
        ' Gold border to indicate current location
        rng.BorderAround xlContinuous, xlMedium, , RGB(201, 162, 39)
        rng.Interior.Color = RGB(60, 50, 20) ' warm highlight
    End If
End Sub

Private Sub ResetMapHighlights()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    
    ' Building ranges to reset
    Dim bldgRanges As Variant
    bldgRanges = Array("L13:N14", "L16:P17", "L19:O20", "L22:O23", _
                       "S13:V14", "S16:V17", "W16:Y17", "W19:Y20", _
                       "S22:V23", "M27:P28", "O9:U9", "O11:U11", _
                       "N4:W4", "L5:N6", "O8:V8", "Q13:R14")
    
    Dim i As Long
    For i = LBound(bldgRanges) To UBound(bldgRanges)
        Dim rng As Range
        Set rng = ws.Range(CStr(bldgRanges(i)))
        rng.Borders(xlEdgeLeft).LineStyle = xlNone
        rng.Borders(xlEdgeRight).LineStyle = xlNone
        rng.Borders(xlEdgeTop).LineStyle = xlNone
        rng.Borders(xlEdgeBottom).LineStyle = xlNone
    Next i
End Sub

'===============================================================
' QUEST PROGRESSION
'===============================================================
Private Sub CheckQuestProgress(sceneID As String)
    Dim wsQ As Worksheet, wsQS As Worksheet
    Set wsQ = ThisWorkbook.Sheets(SH_QUESTS)
    Set wsQS = ThisWorkbook.Sheets(SH_QUESTSTAGES)
    
    Dim r As Long
    For r = 2 To wsQS.Cells(wsQS.Rows.Count, 1).End(xlUp).Row
        Dim questID As String
        questID = CStr(wsQS.Cells(r, 1).Value)
        Dim stageIdx As Long
        stageIdx = CLng(wsQS.Cells(r, 2).Value)
        Dim trigger As String
        trigger = CStr(wsQS.Cells(r, 5).Value)
        Dim triggerType As String
        triggerType = CStr(wsQS.Cells(r, 6).Value)
        
        ' Check if this stage should advance
        Dim shouldAdvance As Boolean
        shouldAdvance = False
        
        If triggerType = "SCENE_COMPLETE" And trigger = sceneID Then
            shouldAdvance = True
        ElseIf triggerType = "FLAG_SET" Then
            shouldAdvance = GetFlag(trigger)
        End If
        
        If shouldAdvance Then
            ' Find quest and check if this is the next stage
            Dim qr As Long
            For qr = 2 To wsQ.Cells(wsQ.Rows.Count, 1).End(xlUp).Row
                If CStr(wsQ.Cells(qr, 1).Value) = questID Then
                    Dim currentStage As Long
                    currentStage = CLng(wsQ.Cells(qr, 6).Value)
                    
                    If stageIdx = currentStage + 1 Or (currentStage = -1 And stageIdx = 0) Then
                        wsQ.Cells(qr, 6).Value = stageIdx
                        wsQ.Cells(qr, 4).Value = "ACTIVE"
                        wsQ.Cells(qr, 5).Value = CStr(wsQS.Cells(r, 4).Value) ' Update description
                        
                        ' Award XP
                        Dim xp As Long
                        Dim xpVal As Variant
                        xpVal = wsQS.Cells(r, 8).Value
                        If IsNumeric(xpVal) Then xp = CLng(xpVal) Else xp = 0
                        If xp > 0 Then
                            Dim curXP As Long
                            curXP = GetStat("XP") + xp
                            SetStat "XP", curXP
                        End If
                    End If
                    Exit For
                End If
            Next qr
        End If
    Next r
End Sub

'===============================================================
' SPECIAL EVENTS
'===============================================================
Private Sub TriggerBlackout()
    ' RAGE hit 100 — trigger blackout scene
    LoadScene "SCN_BLACKOUT"
End Sub

'===============================================================
' VISUAL FEEDBACK
'===============================================================
Private Sub FlashButton(btnNum As Long, success As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GAME)
    
    Dim btn As Shape
    On Error Resume Next
    Set btn = ws.Shapes(BTN_PREFIX & btnNum)
    On Error GoTo 0
    
    If btn Is Nothing Then Exit Sub
    
    If success Then
        btn.Fill.ForeColor.RGB = RGB(40, 50, 30)  ' green flash
    Else
        btn.Fill.ForeColor.RGB = RGB(60, 20, 20)  ' red flash — requirement not met
        btn.TextFrame2.TextRange.Text = btn.TextFrame2.TextRange.Text & "  [LOCKED]"
    End If
    
    ' Reset after brief pause
    Application.Wait Now + TimeValue("00:00:01")
    btn.Fill.ForeColor.RGB = RGB(34, 26, 18)
End Sub

'===============================================================
' SAVE / LOAD SYSTEM
'===============================================================
Public Sub SaveGame(slotNum As Long)
    If slotNum < 1 Or slotNum > 3 Then Exit Sub
    
    Dim wsSave As Worksheet
    Set wsSave = ThisWorkbook.Sheets(SH_SAVES)
    Dim wsGame As Worksheet
    Set wsGame = ThisWorkbook.Sheets(SH_GAME)
    
    Dim r As Long
    r = slotNum + 1
    
    wsSave.Cells(r, 2).Value = "Save " & slotNum & " - Day " & GetStat("DAY_COUNTER")
    wsSave.Cells(r, 3).Value = Now()
    wsSave.Cells(r, 4).Value = wsGame.Range(SCENE_ID_CELL).Value
    wsSave.Cells(r, 5).Value = GetStat("DAY_COUNTER")
    wsSave.Cells(r, 6).Value = GetStat("HEALTH")
    wsSave.Cells(r, 7).Value = GetStat("HUMANITY")
    
    ' Serialize all stats and flags
    Dim serialized As String
    Dim wsStats As Worksheet
    Set wsStats = ThisWorkbook.Sheets(SH_STATS)
    Dim sr As Long
    For sr = 2 To wsStats.Cells(wsStats.Rows.Count, 1).End(xlUp).Row
        serialized = serialized & CStr(wsStats.Cells(sr, 1).Value) & "=" & _
                     CStr(wsStats.Cells(sr, 3).Value) & ";"
    Next sr
    
    serialized = serialized & "|||"
    
    Dim wsFlags As Worksheet
    Set wsFlags = ThisWorkbook.Sheets(SH_FLAGS)
    For sr = 2 To wsFlags.Cells(wsFlags.Rows.Count, 1).End(xlUp).Row
        Dim fv As Variant
        fv = wsFlags.Cells(sr, 2).Value
        If Not IsEmpty(fv) And CStr(fv & "") <> "" Then
            If IsNumeric(fv) Then
                If CLng(fv) <> 0 Then serialized = serialized & CStr(wsFlags.Cells(sr, 1).Value) & ";"
            ElseIf CBool(fv) Then
                serialized = serialized & CStr(wsFlags.Cells(sr, 1).Value) & ";"
            End If
        End If
    Next sr
    
    wsSave.Cells(r, 8).Value = serialized
    
    MsgBox "Game saved to Slot " & slotNum & ".", vbInformation, "SAVED"
End Sub

Public Sub LoadSavedGame(slotNum As Long)
    If slotNum < 1 Or slotNum > 3 Then Exit Sub
    
    Dim wsSave As Worksheet
    Set wsSave = ThisWorkbook.Sheets(SH_SAVES)
    
    Dim r As Long
    r = slotNum + 1
    
    If CStr(wsSave.Cells(r, 2).Value) = "(empty)" Then
        MsgBox "No save in Slot " & slotNum & ".", vbExclamation, "LOAD"
        Exit Sub
    End If
    
    Dim serialized As String
    serialized = CStr(wsSave.Cells(r, 8).Value)
    
    ' Parse stats
    Dim sections() As String
    sections = Split(serialized, "|||")
    
    If UBound(sections) >= 0 Then
        Dim statPairs() As String
        statPairs = Split(sections(0), ";")
        Dim i As Long
        For i = LBound(statPairs) To UBound(statPairs)
            If InStr(statPairs(i), "=") > 0 Then
                Dim kv() As String
                kv = Split(statPairs(i), "=")
                SetStat CStr(kv(0)), kv(1)
            End If
        Next i
    End If
    
    ' Parse flags — reset all first, then set saved ones
    Dim wsFlags As Worksheet
    Set wsFlags = ThisWorkbook.Sheets(SH_FLAGS)
    Dim fr As Long
    For fr = 2 To wsFlags.Cells(wsFlags.Rows.Count, 1).End(xlUp).Row
        wsFlags.Cells(fr, 2).Value = False
    Next fr
    
    If UBound(sections) >= 1 Then
        Dim flagNames() As String
        flagNames = Split(sections(1), ";")
        For i = LBound(flagNames) To UBound(flagNames)
            If Trim(flagNames(i)) <> "" Then
                SetFlag Trim(flagNames(i)), True
            End If
        Next i
    End If
    
    ' Load scene
    Dim sceneID As String
    sceneID = CStr(wsSave.Cells(r, 4).Value)
    LoadScene sceneID
    
    MsgBox "Game loaded from Slot " & slotNum & ".", vbInformation, "LOADED"
End Sub

'===============================================================
' UTILITY — Find scene row
'===============================================================
Private Function FindSceneRow(sceneID As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_SCENES)
    
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = sceneID Then
            FindSceneRow = r
            Exit Function
        End If
    Next r
    FindSceneRow = 0
End Function

'===============================================================
' QUICK SAVE/LOAD MACROS (assign to buttons if desired)
'===============================================================
Public Sub QuickSave1(): SaveGame 1: End Sub
Public Sub QuickSave2(): SaveGame 2: End Sub
Public Sub QuickSave3(): SaveGame 3: End Sub
Public Sub QuickLoad1(): LoadSavedGame 1: End Sub
Public Sub QuickLoad2(): LoadSavedGame 2: End Sub
Public Sub QuickLoad3(): LoadSavedGame 3: End Sub

'===============================================================
' NEW GAME
'===============================================================
Public Sub NewGame()
    If MsgBox("Start a new game? All progress will be lost.", _
              vbYesNo + vbExclamation, "NEW GAME") = vbNo Then Exit Sub
    
    ' Reset all stats to base values
    Dim wsStats As Worksheet
    Set wsStats = ThisWorkbook.Sheets(SH_STATS)
    Dim r As Long
    For r = 2 To wsStats.Cells(wsStats.Rows.Count, 1).End(xlUp).Row
        wsStats.Cells(r, 3).Value = wsStats.Cells(r, 2).Value
    Next r
    
    ' Reset all flags
    Dim wsFlags As Worksheet
    Set wsFlags = ThisWorkbook.Sheets(SH_FLAGS)
    For r = 2 To wsFlags.Cells(wsFlags.Rows.Count, 1).End(xlUp).Row
        wsFlags.Cells(r, 2).Value = False
    Next r
    
    ' Clear inventory
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets(SH_INV)
    For r = 2 To wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
        wsInv.Cells(r, 2).Value = ""
        wsInv.Cells(r, 3).Value = ""
        wsInv.Cells(r, 4).Value = 0
        wsInv.Cells(r, 5).Value = False
    Next r
    
    ' Reset quests
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Sheets(SH_QUESTS)
    For r = 2 To wsQ.Cells(wsQ.Rows.Count, 1).End(xlUp).Row
        Dim qType As String
        qType = CStr(wsQ.Cells(r, 3).Value)
        If qType = "MAIN" Then
            wsQ.Cells(r, 4).Value = "ACTIVE"
            wsQ.Cells(r, 6).Value = 0
        Else
            wsQ.Cells(r, 4).Value = "INACTIVE"
            wsQ.Cells(r, 6).Value = -1
        End If
    Next r
    
    ' Load starting scene
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(SH_CONFIG)
    Dim startScene As String
    startScene = "SCN_PROLOGUE"
    For r = 2 To wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
        If CStr(wsConfig.Cells(r, 1).Value) = "StartingScene" Then
            startScene = CStr(wsConfig.Cells(r, 2).Value)
            Exit For
        End If
    Next r
    
    LoadScene startScene
End Sub
