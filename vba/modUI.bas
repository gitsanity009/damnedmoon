Attribute VB_Name = "modUI"
'===============================================================
' modUI — Game Screen Rendering
' Damned Moon VBA RPG Engine — Phase 2
'===============================================================
' All visual updates to the Game sheet flow through here.
' Narrative display, choice buttons, stats HUD, quest panel,
' inventory panel, day/time/moon display, and map highlighting.
'===============================================================

Option Explicit

'===============================================================
' NARRATIVE — Show story text on the Game sheet
'===============================================================
Public Sub ShowNarrative(text As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    ws.Range(modConfig.NARRATIVE_CELL).Value = text
End Sub

'===============================================================
' CHOICE BUTTONS — Show / Hide / Flash
'===============================================================

' Show a choice button with text and availability state
Public Sub ShowChoiceButton(btnNum As Long, displayText As String, isAvailable As Boolean)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    Dim btn As Shape
    Set btn = GetButton(ws, btnNum)
    If btn Is Nothing Then Exit Sub

    btn.Visible = msoTrue
    btn.TextFrame2.TextRange.text = displayText

    If isAvailable Then
        ' Gold text, dark panel
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(modConfig.C_GOLD_R, modConfig.C_GOLD_G, modConfig.C_GOLD_B)
        btn.Fill.ForeColor.RGB = _
            RGB(modConfig.C_PANEL_R, modConfig.C_PANEL_G, modConfig.C_PANEL_B)
    Else
        ' Dimmed text, darker panel (locked)
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(modConfig.C_DIM_R, modConfig.C_DIM_G, modConfig.C_DIM_B)
        btn.Fill.ForeColor.RGB = _
            RGB(modConfig.C_LOCKED_R, modConfig.C_LOCKED_G, modConfig.C_LOCKED_B)
    End If
End Sub

' Hide a choice button
Public Sub HideChoiceButton(btnNum As Long)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    Dim btn As Shape
    Set btn = GetButton(ws, btnNum)
    If btn Is Nothing Then Exit Sub

    btn.Visible = msoFalse
End Sub

' Flash a button to indicate success or failure
Public Sub FlashButton(btnNum As Long, success As Boolean)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    Dim btn As Shape
    Set btn = GetButton(ws, btnNum)
    If btn Is Nothing Then Exit Sub

    If success Then
        btn.Fill.ForeColor.RGB = RGB(40, 50, 30)  ' green flash
    Else
        btn.Fill.ForeColor.RGB = RGB(60, 20, 20)  ' red flash
        btn.TextFrame2.TextRange.text = btn.TextFrame2.TextRange.text & "  [LOCKED]"
    End If

    ' Brief pause then reset
    Application.Wait Now + TimeValue("00:00:01")
    btn.Fill.ForeColor.RGB = _
        RGB(modConfig.C_PANEL_R, modConfig.C_PANEL_G, modConfig.C_PANEL_B)
End Sub

' Get a button shape by number (returns Nothing if not found)
Private Function GetButton(ws As Worksheet, btnNum As Long) As Shape
    On Error Resume Next
    Set GetButton = ws.Shapes(modConfig.BTN_PREFIX & btnNum)
    On Error GoTo 0
End Function

'===============================================================
' STATS PANEL — Update stat values on Game sheet
'===============================================================
Public Sub UpdateStatsPanel()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    ' Core stats display in rows 7-12, column F
    Dim stats As Variant
    stats = Split(modConfig.CORE_STATS, ",")

    Dim i As Long
    For i = LBound(stats) To UBound(stats)
        Dim statName As String
        statName = Trim(CStr(stats(i)))
        Dim val As Long
        val = modState.GetStat(statName)
        ws.Cells(7 + i, 6).Value = val  ' Column F
    Next i

    ' HP display
    Dim hp As Long
    hp = modState.GetStat(modConfig.STAT_HEALTH)
    ws.Range(modConfig.HP_DISPLAY_CELL).Value = "  HP: " & hp & " / 100"

    ' Day display
    ws.Range(modConfig.DAY_CELL).Value = "DAY " & modState.GetCurrentDay()
    ws.Range(modConfig.TIME_CELL).Value = modState.GetTimeOfDay()

    ' Moon phase
    UpdateMoonDisplay
End Sub

'===============================================================
' MOON DISPLAY
'===============================================================
Private Sub UpdateMoonDisplay()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    Dim wsMoon As Worksheet
    Set wsMoon = modConfig.GetSheet(modConfig.SH_MOON)
    If wsMoon Is Nothing Then Exit Sub

    Dim day As Long
    day = modState.GetCurrentDay()

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsMoon, 1)
        Dim dayRange As String
        dayRange = modUtils.SafeStr(wsMoon.Cells(r, 2).Value)

        If InStr(dayRange, "-") > 0 Then
            Dim parts() As String
            parts = Split(dayRange, "-")
            If day >= modUtils.SafeLng(parts(0), 0) And day <= modUtils.SafeLng(parts(1), 0) Then
                Dim phase As String
                phase = modUtils.SafeStr(wsMoon.Cells(r, 1).Value)
                ws.Range(modConfig.MOON_CELL).Value = GetMoonSymbol(phase) & " " & UCase(phase)
                Exit Sub
            End If
        ElseIf IsNumeric(dayRange) Then
            If day = CLng(dayRange) Then
                ws.Range(modConfig.MOON_CELL).Value = ChrW(&H25CF) & " FULL MOON"
                Exit Sub
            End If
        End If
    Next r
End Sub

' Map a moon phase name to its Unicode symbol
Private Function GetMoonSymbol(phase As String) As String
    Select Case phase
        Case "New Moon": GetMoonSymbol = ChrW(&H25CB)
        Case "Waxing Crescent": GetMoonSymbol = ChrW(&H25D1)
        Case "First Quarter": GetMoonSymbol = ChrW(&H25D1)
        Case "Waxing Gibbous": GetMoonSymbol = ChrW(&H25D0)
        Case "Full Moon", "Full Moon (Night 30)": GetMoonSymbol = ChrW(&H25CF)
        Case "Waning Gibbous": GetMoonSymbol = ChrW(&H25D0)
        Case "Last Quarter": GetMoonSymbol = ChrW(&H25D1)
        Case "Waning Crescent": GetMoonSymbol = ChrW(&H25D1)
        Case Else: GetMoonSymbol = ChrW(&H25CB)
    End Select
End Function

'===============================================================
' QUEST PANEL
'===============================================================
Public Sub UpdateQuestPanel()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    Dim wsQ As Worksheet
    Set wsQ = modConfig.GetSheet(modConfig.SH_QUESTS)
    If wsQ Is Nothing Then
        ws.Range(modConfig.QUEST_DISPLAY_CELL).Value = "(No quests)"
        Exit Sub
    End If

    Dim questText As String
    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsQ, 1)
        If modUtils.SafeStr(wsQ.Cells(r, 4).Value) = "ACTIVE" Then
            Dim qName As String
            qName = modUtils.SafeStr(wsQ.Cells(r, 2).Value)
            Dim qDesc As String
            qDesc = modUtils.SafeStr(wsQ.Cells(r, 5).Value)
            questText = questText & ChrW(&H25C6) & " " & qName & vbLf & _
                        "   " & qDesc & vbLf & vbLf
        End If
    Next r

    If Len(questText) = 0 Then questText = "(No active quests)"
    ws.Range(modConfig.QUEST_DISPLAY_CELL).Value = questText
End Sub

'===============================================================
' INVENTORY PANEL
'===============================================================
Public Sub UpdateInventoryPanel()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Sub

    ' Update weapon display
    Dim weaponID As String
    weaponID = modState.GetEquippedWeapon()
    Dim weaponName As String
    If Len(weaponID) > 0 Then
        Dim wRow As Long
        wRow = modData.GetItemRow(weaponID)
        If wRow > 0 Then
            Dim wsItems As Worksheet
            Set wsItems = modConfig.GetSheet(modConfig.SH_ITEMS)
            If Not wsItems Is Nothing Then
                weaponName = modUtils.SafeStr(wsItems.Cells(wRow, 2).Value)
            End If
        End If
    End If
    If Len(weaponName) = 0 Then weaponName = "(none)"
    ws.Range(modConfig.WEAPON_DISPLAY_CELL).Value = ChrW(&H2694) & " Weapon: " & weaponName

    ' Update inventory slots (rows 7-11, column H/I)
    Dim slotIdx As Long
    slotIdx = 0
    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, 1)
        Dim iName As String
        iName = modUtils.SafeStr(wsInv.Cells(r, 3).Value)
        If Len(iName) > 0 Then
            If slotIdx < 5 Then
                ws.Cells(7 + slotIdx, 8).Value = "  " & iName  ' Column H
                Dim qty As Long
                qty = modUtils.SafeLng(wsInv.Cells(r, 4).Value, 0)
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

'===============================================================
' DAY / TIME PANEL
'===============================================================
Public Sub UpdateDayTimePanel()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    ws.Range(modConfig.DAY_CELL).Value = "DAY " & modState.GetCurrentDay()
    ws.Range(modConfig.TIME_CELL).Value = modState.GetTimeOfDay()
End Sub

'===============================================================
' MAP HIGHLIGHTING
'===============================================================
Public Sub UpdateMapHighlight(locationCode As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    ' Update location label
    Dim locLabel As String
    locLabel = GetLocationLabel(locationCode)
    ws.Range(modConfig.MAP_LOCATION_CELL).Value = "Current: " & locLabel

    ' Reset all building highlights
    ResetMapHighlights

    ' Highlight current location
    Dim hlRange As String
    hlRange = GetLocationRange(locationCode)

    If Len(hlRange) > 0 Then
        Dim rng As Range
        Set rng = ws.Range(hlRange)
        rng.BorderAround xlContinuous, xlMedium, , _
            RGB(modConfig.C_GOLD_R, modConfig.C_GOLD_G, modConfig.C_GOLD_B)
        rng.Interior.Color = _
            RGB(modConfig.C_HIGHLIGHT_R, modConfig.C_HIGHLIGHT_G, modConfig.C_HIGHLIGHT_B)
    End If
End Sub

' Map location codes to display names
Private Function GetLocationLabel(locationCode As String) As String
    Select Case locationCode
        Case "LOS_ANGELES": GetLocationLabel = "LOS ANGELES"
        Case "FIELD", "FIELD_EDGE": GetLocationLabel = "THE FIELD"
        Case "ROAD": GetLocationLabel = "NORTH ROAD"
        Case "MAIN_ROAD": GetLocationLabel = "MAIN ROAD"
        Case "INN", "INN_ROOM", "INN_KITCHEN": GetLocationLabel = "HATTIE'S INN"
        Case "FEED_STORE": GetLocationLabel = "CALHOUN'S FEED STORE"
        Case "SHERIFF_OFFICE": GetLocationLabel = "SHERIFF'S OFFICE"
        Case "CHURCH", "CHURCH_OFFICE": GetLocationLabel = "WROUGHTWOOD CHURCH"
        Case "SWAMP", "SWAMP_CABIN": GetLocationLabel = "MARIE'S SWAMP"
        Case "BARN": GetLocationLabel = "THE BARN"
        Case "MARSH_HOUSE": GetLocationLabel = "MARSH HOUSE"
        Case "NORTH_ROAD": GetLocationLabel = "NORTH ROAD"
        Case "NORTH_WOODS", "DEEP_NORTH_WOODS": GetLocationLabel = "NORTH WOODS"
        Case "PACK_CLEARING": GetLocationLabel = "PACK CLEARING"
        Case "TREE_LINE": GetLocationLabel = "TREE LINE"
        Case "MILL": GetLocationLabel = "THE MILL"
        Case Else: GetLocationLabel = UCase(locationCode)
    End Select
End Function

' Map location codes to cell ranges for map highlighting
Private Function GetLocationRange(locationCode As String) As String
    Select Case locationCode
        Case "FIELD", "FIELD_EDGE": GetLocationRange = "O9:U9"
        Case "ROAD", "NORTH_ROAD": GetLocationRange = "O11:U11"
        Case "INN", "INN_ROOM", "INN_KITCHEN": GetLocationRange = "L16:P17"
        Case "FEED_STORE": GetLocationRange = "S13:V14"
        Case "SHERIFF_OFFICE": GetLocationRange = "S16:V17"
        Case "CHURCH", "CHURCH_OFFICE": GetLocationRange = "L13:N14"
        Case "SWAMP", "SWAMP_CABIN": GetLocationRange = "M27:P28"
        Case "BARN": GetLocationRange = "W19:Y20"
        Case "MARSH_HOUSE": GetLocationRange = "L22:O23"
        Case "NORTH_WOODS", "DEEP_NORTH_WOODS": GetLocationRange = "N4:W4"
        Case "PACK_CLEARING": GetLocationRange = "L5:N6"
        Case "TREE_LINE": GetLocationRange = "O8:V8"
        Case "MAIN_ROAD": GetLocationRange = "Q13:R14"
        Case Else: GetLocationRange = ""
    End Select
End Function

' Reset all map building highlights to default
Private Sub ResetMapHighlights()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

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
