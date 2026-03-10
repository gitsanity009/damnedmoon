Attribute VB_Name = "modConfig"
'===============================================================
' modConfig — Central Configuration & Constants
' Damned Moon VBA RPG Engine
'===============================================================
' All sheet names, table names, layout constants, and config
' values live here. No hardcoded references scattered elsewhere.
'===============================================================

Option Explicit

' ── SHEET NAMES ──
Public Const SH_GAME As String = "Game"
Public Const SH_SCENES As String = "tbl_Scenes"
Public Const SH_FLAGS As String = "tbl_Flags"
Public Const SH_STATS As String = "Stats"
Public Const SH_ITEMS As String = "tbl_ItemDB"
Public Const SH_INV As String = "tbl_Inventory"
Public Const SH_QUESTS As String = "tbl_Quests"
Public Const SH_QUESTSTAGES As String = "tbl_QuestStages"
Public Const SH_ENEMIES As String = "tbl_Enemies"
Public Const SH_MOON As String = "tbl_MoonPhases"
Public Const SH_JOBS As String = "tbl_Jobs"
Public Const SH_COMBAT As String = "tbl_CombatLog"
Public Const SH_SAVES As String = "SaveSlots"
Public Const SH_CONFIG As String = "Config"
Public Const SH_MAPNODES As String = "tbl_MapNodes"
Public Const SH_MAPLINKS As String = "tbl_MapLinks"
Public Const SH_NPCS As String = "tbl_NPCs"
Public Const SH_ENCOUNTERS As String = "tbl_Encounters"
Public Const SH_JOURNAL As String = "tbl_JournalEntries"
Public Const SH_ENDINGS As String = "tbl_Endings"

' ── TABLE NAMES (ListObject names inside sheets) ──
Public Const TBL_SCENES As String = "tbl_Scenes"
Public Const TBL_FLAGS As String = "tbl_Flags"
Public Const TBL_STATS As String = "tbl_Stats"
Public Const TBL_ITEMDB As String = "tbl_ItemDB"
Public Const TBL_INVENTORY As String = "tbl_Inventory"
Public Const TBL_QUESTS As String = "tbl_Quests"
Public Const TBL_QUESTSTAGES As String = "tbl_QuestStages"
Public Const TBL_ENEMIES As String = "tbl_Enemies"
Public Const TBL_MOONPHASES As String = "tbl_MoonPhases"
Public Const TBL_JOBS As String = "tbl_Jobs"
Public Const TBL_COMBATLOG As String = "tbl_CombatLog"
Public Const TBL_MAPNODES As String = "tbl_MapNodes"
Public Const TBL_MAPLINKS As String = "tbl_MapLinks"
Public Const TBL_NPCS As String = "tbl_NPCs"
Public Const TBL_ENCOUNTERS As String = "tbl_Encounters"
Public Const TBL_JOURNAL As String = "tbl_JournalEntries"
Public Const TBL_ENDINGS As String = "tbl_Endings"

' ── GAME SHEET LAYOUT CELLS ──
Public Const NARRATIVE_CELL As String = "B6"
Public Const SCENE_ID_CELL As String = "E40"
Public Const CHOICE_COUNT_CELL As String = "E41"
Public Const LOCATION_CELL As String = "E42"
Public Const DAY_CELL As String = "E2"
Public Const TIME_CELL As String = "E3"
Public Const MOON_CELL As String = "H2"
Public Const MAP_LOCATION_CELL As String = "L3"
Public Const HP_DISPLAY_CELL As String = "E15"
Public Const QUEST_DISPLAY_CELL As String = "E18"
Public Const WEAPON_DISPLAY_CELL As String = "H6"

' ── CHOICE LAYOUT ──
Public Const CHOICE_START_ROW As Long = 25
Public Const CHOICE_END_ROW As Long = 29
Public Const MAX_CHOICES As Long = 5
Public Const CHOICE_COL_SPAN As Long = 4       ' columns per choice block (text, target, req, effect)
Public Const CHOICE_BASE_COL As Long = 7        ' column G = first choice text

' ── SCENE TABLE COLUMNS ──
Public Const SCN_COL_ID As Long = 1             ' A: SceneID
Public Const SCN_COL_NAME As Long = 2           ' B: Scene name
Public Const SCN_COL_LOCATION As Long = 3       ' C: Location code
Public Const SCN_COL_DAY As Long = 4            ' D: Day range
Public Const SCN_COL_TIME As Long = 5           ' E: Time slot
Public Const SCN_COL_NARRATIVE As Long = 6      ' F: Story text
Public Const SCN_COL_ONENTER As Long = 27       ' AA: OnEnter effects
Public Const SCN_COL_ONEXIT As Long = 28        ' AB: OnExit effects
Public Const SCN_COL_COMBAT As Long = 29        ' AC: Combat enemy ID

' ── BUTTON NAMES ──
Public Const BTN_PREFIX As String = "btnChoice"

' ── STAT NAMES ──
Public Const STAT_HEALTH As String = "HEALTH"
Public Const STAT_HUMANITY As String = "HUMANITY"
Public Const STAT_RAGE As String = "RAGE"
Public Const STAT_HUNGER As String = "HUNGER"
Public Const STAT_COMPOSURE As String = "COMPOSURE"
Public Const STAT_INSTINCT As String = "INSTINCT"
Public Const STAT_DAY_COUNTER As String = "DAY_COUNTER"
Public Const STAT_TIME_OF_DAY As String = "TIME_OF_DAY"
Public Const STAT_MOON_PHASE As String = "MOON_PHASE"
Public Const STAT_XP As String = "XP"
Public Const STAT_MONEY As String = "MONEY"

' ── CORE STAT LIST (for clamping, display) ──
Public Const CORE_STATS As String = "HEALTH,HUMANITY,RAGE,HUNGER,COMPOSURE,INSTINCT"

' ── SAVE SYSTEM ──
Public Const SAVE_SLOT_COUNT As Long = 3

' ── UI COLORS (RGB) ──
Public Const C_GOLD_R As Long = 201
Public Const C_GOLD_G As Long = 162
Public Const C_GOLD_B As Long = 39

Public Const C_PANEL_R As Long = 34
Public Const C_PANEL_G As Long = 26
Public Const C_PANEL_B As Long = 18

Public Const C_BORDER_R As Long = 58
Public Const C_BORDER_G As Long = 46
Public Const C_BORDER_B As Long = 34

Public Const C_DIM_R As Long = 80
Public Const C_DIM_G As Long = 70
Public Const C_DIM_B As Long = 60

Public Const C_LOCKED_R As Long = 20
Public Const C_LOCKED_G As Long = 16
Public Const C_LOCKED_B As Long = 12

Public Const C_HIGHLIGHT_R As Long = 60
Public Const C_HIGHLIGHT_G As Long = 50
Public Const C_HIGHLIGHT_B As Long = 20

' ── EFFECT / REQUIREMENT DELIMITERS ──
Public Const EFFECT_DELIM As String = "|"
Public Const TOKEN_DELIM As String = ":"
Public Const SAVE_STAT_DELIM As String = ";"
Public Const SAVE_SECTION_DELIM As String = "|||"

' ── DEFAULT VALUES ──
Public Const DEFAULT_START_SCENE As String = "SCN_PROLOGUE"
Public Const DEFAULT_START_LOCATION As String = "FIELD"

' ── DEBUG ──
Public Const DEBUG_MODE As Boolean = False

'===============================================================
' UTILITY GETTERS
'===============================================================

' Safe worksheet getter — returns Nothing if sheet doesn't exist
Public Function GetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
End Function

' Safe ListObject (table) getter — returns Nothing if not found
Public Function GetTable(sheetName As String, tableName As String) As ListObject
    Dim ws As Worksheet
    Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

' Read a config value from the Config sheet by key name
Public Function GetConfigValue(key As String, Optional fallback As String = "") As String
    Dim ws As Worksheet
    Set ws = GetSheet(SH_CONFIG)
    If ws Is Nothing Then
        GetConfigValue = fallback
        Exit Function
    End If

    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = key Then
            GetConfigValue = CStr(ws.Cells(r, 2).Value & "")
            Exit Function
        End If
    Next r

    GetConfigValue = fallback
End Function
