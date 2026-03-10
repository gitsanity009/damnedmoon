Attribute VB_Name = "modJournal"
'===============================================================
' modJournal — Player Journal & Log System
' Damned Moon VBA RPG Engine — Phase 5
'===============================================================
' Manages journal entries that record quest progress, lore
' discoveries, and personal notes. Entries are timestamped
' with in-game day/time and optionally linked to quests.
'
' tbl_JournalEntries columns:
'   A: EntryID   B: Day   C: Time   D: QuestID (optional)
'   E: Text      F: Type (QUEST, LORE, PERSONAL, COMBAT, SYSTEM)
'   G: Timestamp (real-world, for sorting)
'===============================================================

Option Explicit

' ── JOURNAL TABLE COLUMNS ──
Private Const JRN_COL_ID As Long = 1
Private Const JRN_COL_DAY As Long = 2
Private Const JRN_COL_TIME As Long = 3
Private Const JRN_COL_QUEST As Long = 4
Private Const JRN_COL_TEXT As Long = 5
Private Const JRN_COL_TYPE As Long = 6
Private Const JRN_COL_TIMESTAMP As Long = 7

' ── ENTRY TYPE CONSTANTS ──
Public Const JTYPE_QUEST As String = "QUEST"
Public Const JTYPE_LORE As String = "LORE"
Public Const JTYPE_PERSONAL As String = "PERSONAL"
Public Const JTYPE_COMBAT As String = "COMBAT"
Public Const JTYPE_SYSTEM As String = "SYSTEM"

' ── AUTO-INCREMENT COUNTER ──
Private mNextEntryNum As Long

'===============================================================
' PUBLIC — Add Journal Entry
'===============================================================

' Add a new journal entry with the current in-game timestamp.
' Returns the generated EntryID, or "" on failure.
Public Function AddEntry(entryText As String, _
                         entryType As String, _
                         Optional questID As String = "") As String
    AddEntry = ""

    If Len(Trim(entryText)) = 0 Then Exit Function

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then
        modUtils.DebugLog "modJournal.AddEntry: journal sheet not found"
        Exit Function
    End If

    ' Generate unique entry ID
    Dim entryID As String
    entryID = GenerateEntryID()

    ' Find next empty row
    Dim nextRow As Long
    nextRow = modUtils.GetLastRow(ws, JRN_COL_ID) + 1

    ' Write entry
    ws.Cells(nextRow, JRN_COL_ID).Value = entryID
    ws.Cells(nextRow, JRN_COL_DAY).Value = modState.GetCurrentDay()
    ws.Cells(nextRow, JRN_COL_TIME).Value = modState.GetTimeOfDay()
    ws.Cells(nextRow, JRN_COL_QUEST).Value = questID
    ws.Cells(nextRow, JRN_COL_TEXT).Value = entryText
    ws.Cells(nextRow, JRN_COL_TYPE).Value = UCase(entryType)
    ws.Cells(nextRow, JRN_COL_TIMESTAMP).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")

    ' Update cache so future lookups find this entry
    If Not modData.gJournalRows Is Nothing Then
        If Not modData.gJournalRows.Exists(entryID) Then
            modData.gJournalRows.Add entryID, nextRow
        End If
    End If

    AddEntry = entryID
    modUtils.DebugLog "modJournal.AddEntry: " & entryID & " [" & entryType & "] " & Left(entryText, 40)
End Function

'===============================================================
' PUBLIC — Add typed entries (convenience wrappers)
'===============================================================

' Add a quest-related journal entry
Public Function AddQuestEntry(questID As String, entryText As String) As String
    AddQuestEntry = AddEntry(entryText, JTYPE_QUEST, questID)
End Function

' Add a lore discovery entry
Public Function AddLoreEntry(entryText As String) As String
    AddLoreEntry = AddEntry(entryText, JTYPE_LORE)
End Function

' Add a personal/diary entry
Public Function AddPersonalEntry(entryText As String) As String
    AddPersonalEntry = AddEntry(entryText, JTYPE_PERSONAL)
End Function

' Add a combat log entry
Public Function AddCombatEntry(entryText As String) As String
    AddCombatEntry = AddEntry(entryText, JTYPE_COMBAT)
End Function

' Add a system note
Public Function AddSystemEntry(entryText As String) As String
    AddSystemEntry = AddEntry(entryText, JTYPE_SYSTEM)
End Function

'===============================================================
' PUBLIC — Read Journal Entries
'===============================================================

' Get a single entry's text by EntryID
Public Function GetEntryText(entryID As String) As String
    GetEntryText = ""
    Dim row As Long
    row = modData.GetJournalRow(entryID)
    If row = 0 Then Exit Function
    GetEntryText = modData.ReadCellStr(modConfig.SH_JOURNAL, row, JRN_COL_TEXT)
End Function

' Get all entries for a specific quest, newest first.
' Returns a formatted string for display.
Public Function GetQuestJournal(questID As String) As String
    GetQuestJournal = ""

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Function

    Dim result As String
    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, JRN_COL_ID)

    ' Read newest first
    Dim r As Long
    For r = lastRow To 2 Step -1
        Dim qID As String
        qID = modUtils.SafeStr(ws.Cells(r, JRN_COL_QUEST).Value)
        If qID = questID Then
            Dim day As Long
            day = modUtils.SafeLng(ws.Cells(r, JRN_COL_DAY).Value, 0)
            Dim tod As String
            tod = modUtils.SafeStr(ws.Cells(r, JRN_COL_TIME).Value)
            Dim txt As String
            txt = modUtils.SafeStr(ws.Cells(r, JRN_COL_TEXT).Value)

            If Len(result) > 0 Then result = result & vbLf & vbLf
            result = result & "[Day " & day & ", " & tod & "]" & vbLf & txt
        End If
    Next r

    GetQuestJournal = result
End Function

' Get all entries of a given type, newest first.
Public Function GetEntriesByType(entryType As String) As String
    GetEntriesByType = ""

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Function

    Dim result As String
    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, JRN_COL_ID)

    Dim r As Long
    For r = lastRow To 2 Step -1
        Dim eType As String
        eType = UCase(modUtils.SafeStr(ws.Cells(r, JRN_COL_TYPE).Value))
        If eType = UCase(entryType) Then
            Dim day As Long
            day = modUtils.SafeLng(ws.Cells(r, JRN_COL_DAY).Value, 0)
            Dim tod As String
            tod = modUtils.SafeStr(ws.Cells(r, JRN_COL_TIME).Value)
            Dim txt As String
            txt = modUtils.SafeStr(ws.Cells(r, JRN_COL_TEXT).Value)

            If Len(result) > 0 Then result = result & vbLf & vbLf
            result = result & "[Day " & day & ", " & tod & "]" & vbLf & txt
        End If
    Next r

    GetEntriesByType = result
End Function

' Get all recent entries (up to maxEntries), newest first.
Public Function GetRecentEntries(Optional maxEntries As Long = 10) As String
    GetRecentEntries = ""

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Function

    Dim result As String
    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, JRN_COL_ID)

    Dim count As Long
    count = 0

    Dim r As Long
    For r = lastRow To 2 Step -1
        Dim eID As String
        eID = modUtils.SafeStr(ws.Cells(r, JRN_COL_ID).Value)
        If Len(eID) = 0 Then GoTo NextEntry

        Dim day As Long
        day = modUtils.SafeLng(ws.Cells(r, JRN_COL_DAY).Value, 0)
        Dim tod As String
        tod = modUtils.SafeStr(ws.Cells(r, JRN_COL_TIME).Value)
        Dim eType As String
        eType = modUtils.SafeStr(ws.Cells(r, JRN_COL_TYPE).Value)
        Dim txt As String
        txt = modUtils.SafeStr(ws.Cells(r, JRN_COL_TEXT).Value)

        If Len(result) > 0 Then result = result & vbLf & vbLf
        result = result & "[Day " & day & ", " & tod & " - " & eType & "]" & vbLf & txt

        count = count + 1
        If count >= maxEntries Then Exit For
NextEntry:
    Next r

    GetRecentEntries = result
End Function

'===============================================================
' PUBLIC — Journal Display (for Game sheet HUD)
'===============================================================

' Update the journal display on the Game sheet.
' Shows the most recent entries in a compact format.
Public Sub UpdateJournalDisplay()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_GAME)
    If ws Is Nothing Then Exit Sub

    Dim journalText As String
    journalText = GetRecentEntries(5)

    If Len(journalText) = 0 Then
        journalText = "(No journal entries yet)"
    End If

    ' Display in a designated journal area on the Game sheet
    ' Using row 33+ for journal display (below choices and warnings)
    ws.Cells(33, 2).Value = ChrW(&H1F4D6) & " JOURNAL"
    ws.Cells(34, 2).Value = journalText
End Sub

'===============================================================
' PUBLIC — Entry Count / Existence
'===============================================================

' Count total journal entries
Public Function GetEntryCount() As Long
    GetEntryCount = 0

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, JRN_COL_ID)
    If lastRow < 2 Then Exit Function

    GetEntryCount = lastRow - 1
End Function

' Count entries of a specific type
Public Function GetTypeCount(entryType As String) As Long
    GetTypeCount = 0

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, JRN_COL_ID)

    Dim r As Long
    For r = 2 To lastRow
        If UCase(modUtils.SafeStr(ws.Cells(r, JRN_COL_TYPE).Value)) = UCase(entryType) Then
            GetTypeCount = GetTypeCount + 1
        End If
    Next r
End Function

' Check if a specific lore entry has been discovered (by checking
' if any entry text contains the keyword).
Public Function HasLoreEntry(keyword As String) As Boolean
    HasLoreEntry = False

    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, JRN_COL_ID)

    Dim r As Long
    For r = 2 To lastRow
        If UCase(modUtils.SafeStr(ws.Cells(r, JRN_COL_TYPE).Value)) = JTYPE_LORE Then
            If InStr(UCase(modUtils.SafeStr(ws.Cells(r, JRN_COL_TEXT).Value)), UCase(keyword)) > 0 Then
                HasLoreEntry = True
                Exit Function
            End If
        End If
    Next r
End Function

'===============================================================
' PUBLIC — Clear Journal (for new game reset)
'===============================================================
Public Sub ClearJournal()
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_JOURNAL)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(ws, JRN_COL_ID)
    If lastRow < 2 Then Exit Sub

    ' Clear all data rows (keep headers)
    Dim r As Long
    For r = 2 To lastRow
        Dim c As Long
        For c = 1 To 7
            ws.Cells(r, c).Value = ""
        Next c
    Next r

    ' Reset the counter
    mNextEntryNum = 0

    ' Clear cache
    If Not modData.gJournalRows Is Nothing Then
        modData.gJournalRows.RemoveAll
    End If

    modUtils.DebugLog "modJournal.ClearJournal: all entries cleared"
End Sub

'===============================================================
' PRIVATE — Entry ID Generation
'===============================================================
Private Function GenerateEntryID() As String
    mNextEntryNum = mNextEntryNum + 1
    GenerateEntryID = "JRN_" & Format(Now, "yyyymmdd") & "_" & Format(mNextEntryNum, "000")
End Function
