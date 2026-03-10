Attribute VB_Name = "modSave"
'===============================================================
' modSave — Save Slots & Persistence System
' Damned Moon VBA RPG Engine — Phase 5
'===============================================================
' Multiple save slots with timestamps, auto-save on scene change,
' stack-based undo/rewind snapshots, and full game reset.
'
' SaveSlots sheet layout:
'   Row 1: Headers
'   Each slot uses a block of rows:
'     Row N+0: SlotNum | Timestamp | SceneID | Location | Day | Time | Moon
'     Row N+1: "STATS" | stat1;val1;stat2;val2;...
'     Row N+2: "FLAGS" | flag1;val1;flag2;val2;...
'     Row N+3: "INVENTORY" | itemID;name;qty;equipped;...
'     Row N+4: "QUESTS" | questID;status;stage;...
'
' Serialization uses SAVE_STAT_DELIM (;) between fields
' and SAVE_SECTION_DELIM (|||) between major sections.
'===============================================================

Option Explicit

' ── ROWS PER SAVE SLOT ──
Private Const ROWS_PER_SLOT As Long = 5

' ── SECTION TAGS ──
Private Const TAG_HEADER As String = "HEADER"
Private Const TAG_STATS As String = "STATS"
Private Const TAG_FLAGS As String = "FLAGS"
Private Const TAG_INVENTORY As String = "INVENTORY"
Private Const TAG_QUESTS As String = "QUESTS"

' ── UNDO STACK (in-memory) ──
Private mUndoStack As Collection
Private Const MAX_UNDO As Long = 10

'===============================================================
' PUBLIC — Save Game to Slot
'===============================================================

' Save current game state to slot 1..SAVE_SLOT_COUNT.
' Returns True on success.
Public Function SaveGame(slotNum As Long) As Boolean
    SaveGame = False

    If slotNum < 1 Or slotNum > modConfig.SAVE_SLOT_COUNT Then
        modUtils.ErrorLog "modSave.SaveGame", "Invalid slot: " & slotNum
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_SAVES)
    If ws Is Nothing Then
        modUtils.ErrorLog "modSave.SaveGame", "SaveSlots sheet not found"
        Exit Function
    End If

    Dim baseRow As Long
    baseRow = 2 + (slotNum - 1) * ROWS_PER_SLOT

    On Error GoTo SaveError

    ' Row 1: Header info
    ws.Cells(baseRow, 1).Value = slotNum
    ws.Cells(baseRow, 2).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(baseRow, 3).Value = modState.GetCurrentScene()
    ws.Cells(baseRow, 4).Value = modState.GetCurrentLocation()
    ws.Cells(baseRow, 5).Value = modState.GetCurrentDay()
    ws.Cells(baseRow, 6).Value = modState.GetTimeOfDay()
    ws.Cells(baseRow, 7).Value = modState.GetMoonPhase()

    ' Row 2: Stats
    ws.Cells(baseRow + 1, 1).Value = TAG_STATS
    ws.Cells(baseRow + 1, 2).Value = SerializeStats()

    ' Row 3: Flags
    ws.Cells(baseRow + 2, 1).Value = TAG_FLAGS
    ws.Cells(baseRow + 2, 2).Value = SerializeFlags()

    ' Row 4: Inventory
    ws.Cells(baseRow + 3, 1).Value = TAG_INVENTORY
    ws.Cells(baseRow + 3, 2).Value = SerializeInventory()

    ' Row 5: Quests
    ws.Cells(baseRow + 4, 1).Value = TAG_QUESTS
    ws.Cells(baseRow + 4, 2).Value = SerializeQuests()

    SaveGame = True
    modUtils.DebugLog "modSave.SaveGame: saved to slot " & slotNum & " scene=" & modState.GetCurrentScene()
    Exit Function

SaveError:
    modUtils.ErrorLog "modSave.SaveGame", "Error saving slot " & slotNum & ": " & Err.Description
End Function

'===============================================================
' PUBLIC — Load Game from Slot
'===============================================================

' Restore game state from a save slot. Returns True on success.
Public Function LoadGame(slotNum As Long) As Boolean
    LoadGame = False

    If slotNum < 1 Or slotNum > modConfig.SAVE_SLOT_COUNT Then
        modUtils.ErrorLog "modSave.LoadGame", "Invalid slot: " & slotNum
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_SAVES)
    If ws Is Nothing Then
        modUtils.ErrorLog "modSave.LoadGame", "SaveSlots sheet not found"
        Exit Function
    End If

    Dim baseRow As Long
    baseRow = 2 + (slotNum - 1) * ROWS_PER_SLOT

    ' Check slot is populated
    Dim savedScene As String
    savedScene = modUtils.SafeStr(ws.Cells(baseRow, 3).Value)
    If Len(savedScene) = 0 Then
        modUtils.DebugLog "modSave.LoadGame: slot " & slotNum & " is empty"
        Exit Function
    End If

    On Error GoTo LoadError

    ' Restore header state
    modState.SetCurrentScene savedScene
    modState.SetCurrentLocation modUtils.SafeStr(ws.Cells(baseRow, 4).Value)
    modState.SetCurrentDay modUtils.SafeLng(ws.Cells(baseRow, 5).Value, 1)
    modState.SetTimeOfDay modUtils.SafeStr(ws.Cells(baseRow, 6).Value)
    modState.SetStat modConfig.STAT_MOON_PHASE, modUtils.SafeStr(ws.Cells(baseRow, 7).Value)

    ' Restore stats
    Dim statsStr As String
    statsStr = modUtils.SafeStr(ws.Cells(baseRow + 1, 2).Value)
    If Len(statsStr) > 0 Then DeserializeStats statsStr

    ' Restore flags
    Dim flagsStr As String
    flagsStr = modUtils.SafeStr(ws.Cells(baseRow + 2, 2).Value)
    If Len(flagsStr) > 0 Then DeserializeFlags flagsStr

    ' Restore inventory
    Dim invStr As String
    invStr = modUtils.SafeStr(ws.Cells(baseRow + 3, 2).Value)
    If Len(invStr) > 0 Then DeserializeInventory invStr

    ' Restore quests
    Dim questStr As String
    questStr = modUtils.SafeStr(ws.Cells(baseRow + 4, 2).Value)
    If Len(questStr) > 0 Then DeserializeQuests questStr

    ' Reload the saved scene
    modSceneEngine.LoadScene savedScene

    LoadGame = True
    modUtils.DebugLog "modSave.LoadGame: loaded slot " & slotNum & " scene=" & savedScene
    Exit Function

LoadError:
    modUtils.ErrorLog "modSave.LoadGame", "Error loading slot " & slotNum & ": " & Err.Description
End Function

'===============================================================
' PUBLIC — Auto-Save (called on scene transitions)
'===============================================================

' Auto-save to a dedicated slot (slot 0 internally, stored after
' the regular slots).
Public Sub AutoSave()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_SAVES)
    If ws Is Nothing Then Exit Sub

    ' Auto-save uses the row block after all regular slots
    Dim baseRow As Long
    baseRow = 2 + modConfig.SAVE_SLOT_COUNT * ROWS_PER_SLOT

    On Error Resume Next

    ws.Cells(baseRow, 1).Value = 0  ' slot 0 = autosave
    ws.Cells(baseRow, 2).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(baseRow, 3).Value = modState.GetCurrentScene()
    ws.Cells(baseRow, 4).Value = modState.GetCurrentLocation()
    ws.Cells(baseRow, 5).Value = modState.GetCurrentDay()
    ws.Cells(baseRow, 6).Value = modState.GetTimeOfDay()
    ws.Cells(baseRow, 7).Value = modState.GetMoonPhase()

    ws.Cells(baseRow + 1, 1).Value = TAG_STATS
    ws.Cells(baseRow + 1, 2).Value = SerializeStats()

    ws.Cells(baseRow + 2, 1).Value = TAG_FLAGS
    ws.Cells(baseRow + 2, 2).Value = SerializeFlags()

    ws.Cells(baseRow + 3, 1).Value = TAG_INVENTORY
    ws.Cells(baseRow + 3, 2).Value = SerializeInventory()

    ws.Cells(baseRow + 4, 1).Value = TAG_QUESTS
    ws.Cells(baseRow + 4, 2).Value = SerializeQuests()

    On Error GoTo 0

    modUtils.DebugLog "modSave.AutoSave: saved at " & modState.GetCurrentScene()
End Sub

' Load the auto-save. Returns True on success.
Public Function LoadAutoSave() As Boolean
    LoadAutoSave = False

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_SAVES)
    If ws Is Nothing Then Exit Function

    Dim baseRow As Long
    baseRow = 2 + modConfig.SAVE_SLOT_COUNT * ROWS_PER_SLOT

    Dim savedScene As String
    savedScene = modUtils.SafeStr(ws.Cells(baseRow, 3).Value)
    If Len(savedScene) = 0 Then Exit Function

    On Error GoTo AutoLoadErr

    modState.SetCurrentScene savedScene
    modState.SetCurrentLocation modUtils.SafeStr(ws.Cells(baseRow, 4).Value)
    modState.SetCurrentDay modUtils.SafeLng(ws.Cells(baseRow, 5).Value, 1)
    modState.SetTimeOfDay modUtils.SafeStr(ws.Cells(baseRow, 6).Value)
    modState.SetStat modConfig.STAT_MOON_PHASE, modUtils.SafeStr(ws.Cells(baseRow, 7).Value)

    Dim statsStr As String: statsStr = modUtils.SafeStr(ws.Cells(baseRow + 1, 2).Value)
    If Len(statsStr) > 0 Then DeserializeStats statsStr

    Dim flagsStr As String: flagsStr = modUtils.SafeStr(ws.Cells(baseRow + 2, 2).Value)
    If Len(flagsStr) > 0 Then DeserializeFlags flagsStr

    Dim invStr As String: invStr = modUtils.SafeStr(ws.Cells(baseRow + 3, 2).Value)
    If Len(invStr) > 0 Then DeserializeInventory invStr

    Dim questStr As String: questStr = modUtils.SafeStr(ws.Cells(baseRow + 4, 2).Value)
    If Len(questStr) > 0 Then DeserializeQuests questStr

    modSceneEngine.LoadScene savedScene
    LoadAutoSave = True
    Exit Function

AutoLoadErr:
    modUtils.ErrorLog "modSave.LoadAutoSave", Err.Description
End Function

'===============================================================
' PUBLIC — Get Save Slot Info (for UI display)
'===============================================================

' Returns a summary string for a slot, or "" if empty.
Public Function GetSlotInfo(slotNum As Long) As String
    GetSlotInfo = ""

    If slotNum < 1 Or slotNum > modConfig.SAVE_SLOT_COUNT Then Exit Function

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_SAVES)
    If ws Is Nothing Then Exit Function

    Dim baseRow As Long
    baseRow = 2 + (slotNum - 1) * ROWS_PER_SLOT

    Dim savedScene As String
    savedScene = modUtils.SafeStr(ws.Cells(baseRow, 3).Value)
    If Len(savedScene) = 0 Then Exit Function

    Dim ts As String
    ts = modUtils.SafeStr(ws.Cells(baseRow, 2).Value)
    Dim loc As String
    loc = modUtils.SafeStr(ws.Cells(baseRow, 4).Value)
    Dim day As String
    day = modUtils.SafeStr(ws.Cells(baseRow, 5).Value)

    GetSlotInfo = "Slot " & slotNum & ": " & savedScene & _
                  " | Day " & day & " | " & loc & " | " & ts
End Function

' Check if a slot has data
Public Function IsSlotOccupied(slotNum As Long) As Boolean
    IsSlotOccupied = (Len(GetSlotInfo(slotNum)) > 0)
End Function

'===============================================================
' PUBLIC — Delete Save Slot
'===============================================================
Public Sub DeleteSave(slotNum As Long)
    If slotNum < 1 Or slotNum > modConfig.SAVE_SLOT_COUNT Then Exit Sub

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_SAVES)
    If ws Is Nothing Then Exit Sub

    Dim baseRow As Long
    baseRow = 2 + (slotNum - 1) * ROWS_PER_SLOT

    Dim r As Long
    For r = baseRow To baseRow + ROWS_PER_SLOT - 1
        Dim c As Long
        For c = 1 To 7
            ws.Cells(r, c).Value = ""
        Next c
    Next r

    modUtils.DebugLog "modSave.DeleteSave: cleared slot " & slotNum
End Sub

'===============================================================
' PUBLIC — Undo Stack (snapshot-based rewind)
'===============================================================

' Push a state snapshot onto the undo stack.
' Call before processing choices for rewind support.
Public Sub PushSnapshot()
    If mUndoStack Is Nothing Then Set mUndoStack = New Collection

    Dim snapshot As String
    snapshot = modState.GetCurrentScene() & modConfig.SAVE_SECTION_DELIM & _
               SerializeStats() & modConfig.SAVE_SECTION_DELIM & _
               SerializeFlags() & modConfig.SAVE_SECTION_DELIM & _
               SerializeInventory() & modConfig.SAVE_SECTION_DELIM & _
               SerializeQuests() & modConfig.SAVE_SECTION_DELIM & _
               modState.GetCurrentLocation() & modConfig.SAVE_SECTION_DELIM & _
               CStr(modState.GetCurrentDay()) & modConfig.SAVE_SECTION_DELIM & _
               modState.GetTimeOfDay() & modConfig.SAVE_SECTION_DELIM & _
               modState.GetMoonPhase()

    ' Trim stack if at max
    Do While mUndoStack.Count >= MAX_UNDO
        mUndoStack.Remove 1
    Loop

    mUndoStack.Add snapshot
    modUtils.DebugLog "modSave.PushSnapshot: stack size=" & mUndoStack.Count
End Sub

' Pop and restore the most recent snapshot. Returns True if restored.
Public Function PopSnapshot() As Boolean
    PopSnapshot = False

    If mUndoStack Is Nothing Then Exit Function
    If mUndoStack.Count = 0 Then Exit Function

    Dim snapshot As String
    snapshot = CStr(mUndoStack(mUndoStack.Count))
    mUndoStack.Remove mUndoStack.Count

    ' Parse sections
    Dim sections As Variant
    sections = Split(snapshot, modConfig.SAVE_SECTION_DELIM)

    If UBound(sections) < 8 Then
        modUtils.ErrorLog "modSave.PopSnapshot", "Corrupt snapshot, not enough sections"
        Exit Function
    End If

    ' Restore state
    Dim sceneID As String: sceneID = CStr(sections(0))
    modState.SetCurrentScene sceneID

    DeserializeStats CStr(sections(1))
    DeserializeFlags CStr(sections(2))
    DeserializeInventory CStr(sections(3))
    DeserializeQuests CStr(sections(4))

    modState.SetCurrentLocation CStr(sections(5))
    modState.SetCurrentDay modUtils.SafeLng(sections(6), 1)
    modState.SetTimeOfDay CStr(sections(7))
    modState.SetStat modConfig.STAT_MOON_PHASE, CStr(sections(8))

    ' Reload scene
    modSceneEngine.LoadScene sceneID

    PopSnapshot = True
    modUtils.DebugLog "modSave.PopSnapshot: rewound to " & sceneID & ", stack=" & mUndoStack.Count
End Function

' Get undo stack depth
Public Function GetUndoDepth() As Long
    If mUndoStack Is Nothing Then
        GetUndoDepth = 0
    Else
        GetUndoDepth = mUndoStack.Count
    End If
End Function

' Clear the undo stack
Public Sub ClearUndoStack()
    Set mUndoStack = New Collection
    modUtils.DebugLog "modSave.ClearUndoStack: cleared"
End Sub

'===============================================================
' PRIVATE — Serialization Helpers
'===============================================================

' Serialize all stats to "Name;Value;Name;Value;..."
Private Function SerializeStats() As String
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_STATS)
    If ws Is Nothing Then
        SerializeStats = ""
        Exit Function
    End If

    Dim parts() As String
    Dim count As Long
    count = 0

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, 1)

    ReDim parts(1 To lastRow * 2)

    Dim r As Long
    For r = 2 To lastRow
        Dim sName As String
        sName = modUtils.SafeStr(ws.Cells(r, 1).Value)
        If Len(sName) > 0 Then
            count = count + 1
            parts(count) = sName
            count = count + 1
            parts(count) = modUtils.SafeStr(ws.Cells(r, 3).Value)
        End If
    Next r

    If count = 0 Then
        SerializeStats = ""
        Exit Function
    End If

    ReDim Preserve parts(1 To count)
    SerializeStats = Join(parts, modConfig.SAVE_STAT_DELIM)
End Function

' Serialize all flags to "Name;Value;Name;Value;..."
Private Function SerializeFlags() As String
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_FLAGS)
    If ws Is Nothing Then
        SerializeFlags = ""
        Exit Function
    End If

    Dim parts() As String
    Dim count As Long
    count = 0

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, 1)

    ReDim parts(1 To lastRow * 2)

    Dim r As Long
    For r = 2 To lastRow
        Dim fName As String
        fName = modUtils.SafeStr(ws.Cells(r, 1).Value)
        If Len(fName) > 0 Then
            count = count + 1
            parts(count) = fName
            count = count + 1
            If modUtils.SafeBool(ws.Cells(r, 2).Value) Then
                parts(count) = "1"
            Else
                parts(count) = "0"
            End If
        End If
    Next r

    If count = 0 Then
        SerializeFlags = ""
        Exit Function
    End If

    ReDim Preserve parts(1 To count)
    SerializeFlags = Join(parts, modConfig.SAVE_STAT_DELIM)
End Function

' Serialize inventory to "ItemID;Name;Qty;Equipped;..."
Private Function SerializeInventory() As String
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_INV)
    If ws Is Nothing Then
        SerializeInventory = ""
        Exit Function
    End If

    Dim parts() As String
    Dim count As Long
    count = 0

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, 1)

    ReDim parts(1 To lastRow * 4)

    Dim r As Long
    For r = 2 To lastRow
        Dim itemID As String
        itemID = modUtils.SafeStr(ws.Cells(r, 2).Value)
        If Len(itemID) > 0 Then
            count = count + 1: parts(count) = itemID
            count = count + 1: parts(count) = modUtils.SafeStr(ws.Cells(r, 3).Value)
            count = count + 1: parts(count) = CStr(modUtils.SafeLng(ws.Cells(r, 4).Value, 0))
            count = count + 1
            If modUtils.SafeBool(ws.Cells(r, 5).Value) Then
                parts(count) = "1"
            Else
                parts(count) = "0"
            End If
        End If
    Next r

    If count = 0 Then
        SerializeInventory = ""
        Exit Function
    End If

    ReDim Preserve parts(1 To count)
    SerializeInventory = Join(parts, modConfig.SAVE_STAT_DELIM)
End Function

' Serialize quests to "QuestID;Status;Stage;..."
Private Function SerializeQuests() As String
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then
        SerializeQuests = ""
        Exit Function
    End If

    Dim parts() As String
    Dim count As Long
    count = 0

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, 1)

    ReDim parts(1 To lastRow * 3)

    Dim r As Long
    For r = 2 To lastRow
        Dim qID As String
        qID = modUtils.SafeStr(ws.Cells(r, 1).Value)
        If Len(qID) > 0 Then
            count = count + 1: parts(count) = qID
            count = count + 1: parts(count) = modUtils.SafeStr(ws.Cells(r, 4).Value)
            count = count + 1: parts(count) = CStr(modUtils.SafeLng(ws.Cells(r, 6).Value, -1))
        End If
    Next r

    If count = 0 Then
        SerializeQuests = ""
        Exit Function
    End If

    ReDim Preserve parts(1 To count)
    SerializeQuests = Join(parts, modConfig.SAVE_STAT_DELIM)
End Function

'===============================================================
' PRIVATE — Deserialization Helpers
'===============================================================

' Restore stats from serialized string
Private Sub DeserializeStats(statsStr As String)
    Dim parts As Variant
    parts = Split(statsStr, modConfig.SAVE_STAT_DELIM)

    Dim i As Long
    For i = LBound(parts) To UBound(parts) - 1 Step 2
        Dim sName As String
        sName = Trim(CStr(parts(i)))
        Dim sVal As String
        sVal = Trim(CStr(parts(i + 1)))
        If Len(sName) > 0 Then
            If IsNumeric(sVal) Then
                modState.SetStat sName, CLng(sVal)
            Else
                modState.SetStat sName, sVal
            End If
        End If
    Next i
End Sub

' Restore flags from serialized string
Private Sub DeserializeFlags(flagsStr As String)
    Dim parts As Variant
    parts = Split(flagsStr, modConfig.SAVE_STAT_DELIM)

    Dim i As Long
    For i = LBound(parts) To UBound(parts) - 1 Step 2
        Dim fName As String
        fName = Trim(CStr(parts(i)))
        Dim fVal As String
        fVal = Trim(CStr(parts(i + 1)))
        If Len(fName) > 0 Then
            modState.SetFlag fName, (fVal = "1" Or UCase(fVal) = "TRUE")
        End If
    Next i
End Sub

' Restore inventory from serialized string
Private Sub DeserializeInventory(invStr As String)
    ' First clear existing inventory
    modState.ResetInventory

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_INV)
    If ws Is Nothing Then Exit Sub

    Dim parts As Variant
    parts = Split(invStr, modConfig.SAVE_STAT_DELIM)

    Dim slotRow As Long
    slotRow = 2

    Dim i As Long
    For i = LBound(parts) To UBound(parts) - 3 Step 4
        Dim itemID As String: itemID = Trim(CStr(parts(i)))
        Dim itemName As String: itemName = Trim(CStr(parts(i + 1)))
        Dim qty As Long: qty = modUtils.SafeLng(parts(i + 2), 0)
        Dim equipped As Boolean: equipped = (Trim(CStr(parts(i + 3))) = "1")

        If Len(itemID) > 0 And qty > 0 Then
            ws.Cells(slotRow, 2).Value = itemID
            ws.Cells(slotRow, 3).Value = itemName
            ws.Cells(slotRow, 4).Value = qty
            ws.Cells(slotRow, 5).Value = equipped
            slotRow = slotRow + 1
        End If
    Next i
End Sub

' Restore quests from serialized string
Private Sub DeserializeQuests(questStr As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_QUESTS)
    If ws Is Nothing Then Exit Sub

    Dim parts As Variant
    parts = Split(questStr, modConfig.SAVE_STAT_DELIM)

    Dim i As Long
    For i = LBound(parts) To UBound(parts) - 2 Step 3
        Dim qID As String: qID = Trim(CStr(parts(i)))
        Dim qStatus As String: qStatus = Trim(CStr(parts(i + 1)))
        Dim qStage As Long: qStage = modUtils.SafeLng(parts(i + 2), -1)

        If Len(qID) > 0 Then
            Dim qRow As Long
            qRow = modData.GetQuestRow(qID)
            If qRow > 0 Then
                ws.Cells(qRow, 4).Value = qStatus
                ws.Cells(qRow, 6).Value = qStage
            End If
        End If
    Next i
End Sub
