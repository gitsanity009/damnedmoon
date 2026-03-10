Attribute VB_Name = "modCombat"
'===============================================================
' modCombat — Turn-Based Combat System
' Damned Moon VBA RPG Engine — Phase 5
'===============================================================
' Quick combat loop: player picks action -> enemy responds ->
' resolve damage -> check victory/defeat/flee. Reads enemy data
' from tbl_Enemies, logs rounds to tbl_CombatLog.
'
' Enemy table columns (tbl_Enemies):
'   A: EnemyID    B: Name      C: Type       D: BaseHP
'   E: BaseDamage F: Defense   G: Speed      H: CritChance
'   I: DropItems  J: DropXP    K: DropMoney  L: Abilities
'   M: WeakElement N: Description
'
' Combat log columns (tbl_CombatLog):
'   A: LogID  B: Timestamp  C: EnemyID  D: Round
'   E: Actor  F: Action     G: Damage   H: Result
'===============================================================

Option Explicit

' ── ENEMY TABLE COLUMNS ──
Private Const EN_COL_ID As Long = 1
Private Const EN_COL_NAME As Long = 2
Private Const EN_COL_TYPE As Long = 3
Private Const EN_COL_HP As Long = 4
Private Const EN_COL_DMG As Long = 5
Private Const EN_COL_DEF As Long = 6
Private Const EN_COL_SPD As Long = 7
Private Const EN_COL_CRIT As Long = 8
Private Const EN_COL_DROPS As Long = 9
Private Const EN_COL_XP As Long = 10
Private Const EN_COL_MONEY As Long = 11
Private Const EN_COL_ABILITIES As Long = 12
Private Const EN_COL_WEAK As Long = 13
Private Const EN_COL_DESC As Long = 14

' ── COMBAT STATE (module-level for multi-round tracking) ──
Private mInCombat As Boolean
Private mEnemyID As String
Private mEnemyHP As Long
Private mEnemyMaxHP As Long
Private mEnemyDmg As Long
Private mEnemyDef As Long
Private mEnemySpd As Long
Private mEnemyCrit As Long
Private mEnemyName As String
Private mRound As Long
Private mCombatLog As String

' ── PLAYER ACTION CONSTANTS ──
Public Const ACT_ATTACK As String = "ATTACK"
Public Const ACT_DEFEND As String = "DEFEND"
Public Const ACT_ITEM As String = "ITEM"
Public Const ACT_FLEE As String = "FLEE"

' ── COMBAT RESULT CONSTANTS ──
Public Const RESULT_VICTORY As String = "VICTORY"
Public Const RESULT_DEFEAT As String = "DEFEAT"
Public Const RESULT_FLED As String = "FLED"
Public Const RESULT_ONGOING As String = "ONGOING"

'===============================================================
' PUBLIC — Initiate Combat
'===============================================================

' Start combat with an enemy. Returns True if combat was set up.
Public Function InitiateCombat(enemyID As String) As Boolean
    InitiateCombat = False

    If Not modData.EnemyExists(enemyID) Then
        modUtils.ErrorLog "modCombat.InitiateCombat", "Enemy '" & enemyID & "' not found"
        Exit Function
    End If

    Dim row As Long
    row = modData.GetEnemyRow(enemyID)

    ' Load enemy stats
    mEnemyID = enemyID
    mEnemyName = modData.ReadCellStr(modConfig.SH_ENEMIES, row, EN_COL_NAME)
    mEnemyHP = modData.ReadCellLng(modConfig.SH_ENEMIES, row, EN_COL_HP)
    mEnemyMaxHP = mEnemyHP
    mEnemyDmg = modData.ReadCellLng(modConfig.SH_ENEMIES, row, EN_COL_DMG)
    mEnemyDef = modData.ReadCellLng(modConfig.SH_ENEMIES, row, EN_COL_DEF)
    mEnemySpd = modData.ReadCellLng(modConfig.SH_ENEMIES, row, EN_COL_SPD)
    mEnemyCrit = modData.ReadCellLng(modConfig.SH_ENEMIES, row, EN_COL_CRIT)

    If mEnemyHP <= 0 Then mEnemyHP = 30
    If mEnemyMaxHP <= 0 Then mEnemyMaxHP = 30
    If mEnemyDmg <= 0 Then mEnemyDmg = 5

    mRound = 0
    mInCombat = True
    mCombatLog = ""

    AppendLog "Combat begins: " & mEnemyName & " (HP:" & mEnemyHP & " DMG:" & mEnemyDmg & ")"

    modUtils.DebugLog "modCombat.InitiateCombat: " & enemyID & " HP=" & mEnemyHP
    InitiateCombat = True
End Function

'===============================================================
' PUBLIC — Execute one combat round
'===============================================================

' Process one round of combat. Returns the result string:
' ONGOING, VICTORY, DEFEAT, or FLED.
Public Function ExecuteRound(playerAction As String) As String
    If Not mInCombat Then
        ExecuteRound = RESULT_DEFEAT
        Exit Function
    End If

    mRound = mRound + 1
    Dim playerHP As Long
    playerHP = modState.GetStat(modConfig.STAT_HEALTH)

    ' ── PLAYER TURN ──
    Dim playerDmg As Long
    Dim defending As Boolean
    defending = False

    Select Case UCase(playerAction)
        Case ACT_ATTACK
            playerDmg = RollPlayerDamage()
            Dim actualDmg As Long
            actualDmg = modUtils.Clamp(playerDmg - mEnemyDef, 1, 999)
            mEnemyHP = mEnemyHP - actualDmg
            AppendLog "Round " & mRound & ": You attack for " & actualDmg & " damage"
            LogCombatRound "PLAYER", ACT_ATTACK, actualDmg, RESULT_ONGOING

        Case ACT_DEFEND
            defending = True
            AppendLog "Round " & mRound & ": You brace for impact"
            LogCombatRound "PLAYER", ACT_DEFEND, 0, RESULT_ONGOING

        Case ACT_ITEM
            ' Use a healing item if available (consume bandage/potion)
            Dim healed As Long
            healed = UseHealingItem()
            If healed > 0 Then
                AppendLog "Round " & mRound & ": You use an item, healing " & healed & " HP"
            Else
                AppendLog "Round " & mRound & ": No usable items! You lose your turn"
            End If
            LogCombatRound "PLAYER", ACT_ITEM, healed, RESULT_ONGOING

        Case ACT_FLEE
            Dim fleeResult As String
            fleeResult = AttemptFlee()
            If fleeResult = RESULT_FLED Then
                AppendLog "Round " & mRound & ": You flee from combat!"
                LogCombatRound "PLAYER", ACT_FLEE, 0, RESULT_FLED
                EndCombat RESULT_FLED
                ExecuteRound = RESULT_FLED
                Exit Function
            Else
                AppendLog "Round " & mRound & ": You fail to escape!"
                LogCombatRound "PLAYER", ACT_FLEE, 0, RESULT_ONGOING
            End If
    End Select

    ' Check enemy defeated
    If mEnemyHP <= 0 Then
        AppendLog mEnemyName & " is defeated!"
        LogCombatRound "ENEMY", "DEFEATED", 0, RESULT_VICTORY
        EndCombat RESULT_VICTORY
        ExecuteRound = RESULT_VICTORY
        Exit Function
    End If

    ' ── ENEMY TURN ──
    Dim enemyDmg As Long
    enemyDmg = RollEnemyDamage()

    ' Defending halves incoming damage
    If defending Then
        enemyDmg = enemyDmg \ 2
        If enemyDmg < 1 Then enemyDmg = 1
    End If

    ' Apply player defense from composure
    Dim playerDef As Long
    playerDef = modState.GetStat(modConfig.STAT_COMPOSURE) \ 10
    enemyDmg = modUtils.Clamp(enemyDmg - playerDef, 1, 999)

    modState.AddStat modConfig.STAT_HEALTH, -CDbl(enemyDmg)
    AppendLog mEnemyName & " attacks for " & enemyDmg & " damage"
    LogCombatRound "ENEMY", ACT_ATTACK, enemyDmg, RESULT_ONGOING

    ' Rage increases during combat
    modState.AddStat modConfig.STAT_RAGE, 3

    ' Check player defeated
    playerHP = modState.GetStat(modConfig.STAT_HEALTH)
    If playerHP <= 0 Then
        AppendLog "You collapse... defeated."
        LogCombatRound "PLAYER", "DEFEATED", 0, RESULT_DEFEAT
        EndCombat RESULT_DEFEAT
        ExecuteRound = RESULT_DEFEAT
        Exit Function
    End If

    ExecuteRound = RESULT_ONGOING
End Function

'===============================================================
' PUBLIC — Quick Combat (auto-resolve for scene triggers)
'===============================================================

' Run a full combat encounter from a scene trigger (column AC).
' Shows narrative and returns the result.
Public Function QuickCombat(enemyID As String) As String
    QuickCombat = RESULT_DEFEAT

    If Not InitiateCombat(enemyID) Then Exit Function

    ' Build combat intro narrative
    Dim intro As String
    intro = ChrW(&H2694) & " COMBAT: " & mEnemyName & vbLf & vbLf
    Dim row As Long
    row = modData.GetEnemyRow(enemyID)
    Dim desc As String
    desc = modData.ReadCellStr(modConfig.SH_ENEMIES, row, EN_COL_DESC)
    If Len(desc) > 0 Then intro = intro & desc & vbLf & vbLf

    ' Auto-resolve: alternate attack rounds until someone drops
    Dim maxRounds As Long
    maxRounds = 20
    Dim result As String
    result = RESULT_ONGOING

    Dim rnd As Long
    For rnd = 1 To maxRounds
        result = ExecuteRound(ACT_ATTACK)
        If result <> RESULT_ONGOING Then Exit For
    Next rnd

    ' If still ongoing after max rounds, player flees
    If result = RESULT_ONGOING Then
        result = RESULT_FLED
        EndCombat RESULT_FLED
    End If

    ' Show combat summary
    Dim summary As String
    summary = intro & GetCombatLog() & vbLf & vbLf

    Select Case result
        Case RESULT_VICTORY
            summary = summary & ChrW(&H2605) & " VICTORY!"
        Case RESULT_DEFEAT
            summary = summary & ChrW(&H2620) & " DEFEATED..."
        Case RESULT_FLED
            summary = summary & ChrW(&H21B6) & " ESCAPED"
    End Select

    modUI.ShowNarrative summary
    QuickCombat = result
End Function

'===============================================================
' PUBLIC — Combat State Queries
'===============================================================

Public Function IsInCombat() As Boolean
    IsInCombat = mInCombat
End Function

Public Function GetEnemyHP() As Long
    GetEnemyHP = mEnemyHP
End Function

Public Function GetEnemyMaxHP() As Long
    GetEnemyMaxHP = mEnemyMaxHP
End Function

Public Function GetEnemyName() As String
    GetEnemyName = mEnemyName
End Function

Public Function GetCurrentRound() As Long
    GetCurrentRound = mRound
End Function

Public Function GetCombatLog() As String
    GetCombatLog = mCombatLog
End Function

' Get the enemy display name from the table
Public Function GetEnemyDisplayName(enemyID As String) As String
    Dim row As Long
    row = modData.GetEnemyRow(enemyID)
    If row = 0 Then
        GetEnemyDisplayName = enemyID
        Exit Function
    End If
    GetEnemyDisplayName = modData.ReadCellStr(modConfig.SH_ENEMIES, row, EN_COL_NAME)
End Function

'===============================================================
' PRIVATE — Damage Rolls
'===============================================================

' Roll player damage: base from weapon + instinct bonus + crit
Private Function RollPlayerDamage() As Long
    Dim baseDmg As Long
    baseDmg = 8  ' unarmed base

    ' Weapon bonus
    Dim weaponID As String
    weaponID = modState.GetEquippedWeapon()
    If Len(weaponID) > 0 Then
        Dim wRow As Long
        wRow = modData.GetItemRow(weaponID)
        If wRow > 0 Then
            ' Item column for damage bonus (column F = effects)
            Dim effectStr As String
            effectStr = modData.ReadCellStr(modConfig.SH_ITEMS, wRow, 6)
            If InStr(effectStr, "DMG:") > 0 Then
                Dim dmgPart As String
                dmgPart = modUtils.StripPrefix(effectStr, "DMG:")
                Dim dmgVal As Long
                dmgVal = modUtils.SafeLng(Left(dmgPart, InStr(dmgPart & "|", "|") - 1), 0)
                If dmgVal > 0 Then baseDmg = dmgVal
            End If
        End If
    End If

    ' Instinct bonus: +1 per 20 instinct
    baseDmg = baseDmg + modState.GetStat(modConfig.STAT_INSTINCT) \ 20

    ' Randomize: 80-120% of base
    Dim roll As Long
    roll = modUtils.RandBetween(baseDmg * 80, baseDmg * 120) \ 100
    If roll < 1 Then roll = 1

    ' Crit check: 10% base chance
    Dim critRoll As Long
    critRoll = modUtils.RandBetween(1, 100)
    If critRoll <= 10 Then
        roll = roll * 2
        AppendLog "  ** CRITICAL HIT! **"
    End If

    RollPlayerDamage = roll
End Function

' Roll enemy damage: base from stats + randomization + crit
Private Function RollEnemyDamage() As Long
    Dim baseDmg As Long
    baseDmg = mEnemyDmg
    If baseDmg < 1 Then baseDmg = 5

    ' Randomize: 75-125%
    Dim roll As Long
    roll = modUtils.RandBetween(baseDmg * 75, baseDmg * 125) \ 100
    If roll < 1 Then roll = 1

    ' Enemy crit
    Dim critChance As Long
    critChance = mEnemyCrit
    If critChance <= 0 Then critChance = 5

    Dim critRoll As Long
    critRoll = modUtils.RandBetween(1, 100)
    If critRoll <= critChance Then
        roll = roll * 2
        AppendLog "  ** " & mEnemyName & " lands a CRITICAL HIT! **"
    End If

    RollEnemyDamage = roll
End Function

'===============================================================
' PRIVATE — Flee Attempt
'===============================================================

' Attempt to flee. Chance based on player speed (composure) vs enemy speed.
Private Function AttemptFlee() As String
    Dim playerSpeed As Long
    playerSpeed = modState.GetStat(modConfig.STAT_COMPOSURE)

    ' Flee chance: 30 base + (playerSpeed - enemySpeed)
    Dim fleeChance As Long
    fleeChance = modUtils.Clamp(30 + (playerSpeed - mEnemySpd), 10, 80)

    Dim roll As Long
    roll = modUtils.RandBetween(1, 100)

    If roll <= fleeChance Then
        AttemptFlee = RESULT_FLED
    Else
        AttemptFlee = RESULT_ONGOING
    End If
End Function

'===============================================================
' PRIVATE — Healing Item Usage
'===============================================================

' Look for a healing consumable in inventory and use it.
' Returns the amount healed (0 if none found).
Private Function UseHealingItem() As Long
    UseHealingItem = 0

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Function

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, 1)
        Dim itemID As String
        itemID = modUtils.SafeStr(wsInv.Cells(r, 2).Value)
        If Len(itemID) = 0 Then GoTo NextItem

        ' Check if this is a consumable with healing
        Dim iRow As Long
        iRow = modData.GetItemRow(itemID)
        If iRow = 0 Then GoTo NextItem

        Dim iType As String
        iType = UCase(modData.ReadCellStr(modConfig.SH_ITEMS, iRow, 3))
        If iType <> "CONSUMABLE" Then GoTo NextItem

        Dim iEffects As String
        iEffects = modData.ReadCellStr(modConfig.SH_ITEMS, iRow, 6)
        If InStr(UCase(iEffects), "HEALTH") = 0 Then GoTo NextItem

        ' Found a healing consumable — use it
        Dim qty As Long
        qty = modUtils.SafeLng(wsInv.Cells(r, 4).Value, 0) - 1
        If qty <= 0 Then
            wsInv.Cells(r, 2).Value = ""
            wsInv.Cells(r, 3).Value = ""
            wsInv.Cells(r, 4).Value = 0
        Else
            wsInv.Cells(r, 4).Value = qty
        End If

        ' Apply the healing effect
        modEffects.ProcessEffects iEffects
        UseHealingItem = 15  ' default healing amount
        Exit Function

NextItem:
    Next r
End Function

'===============================================================
' PRIVATE — End Combat & Award Rewards
'===============================================================
Private Sub EndCombat(result As String)
    mInCombat = False

    If result = RESULT_VICTORY Then
        Dim row As Long
        row = modData.GetEnemyRow(mEnemyID)
        If row > 0 Then
            ' Award XP
            Dim xp As Long
            xp = modData.ReadCellLng(modConfig.SH_ENEMIES, row, EN_COL_XP)
            If xp > 0 Then modState.AddStat modConfig.STAT_XP, xp

            ' Award money
            Dim money As Long
            money = modData.ReadCellLng(modConfig.SH_ENEMIES, row, EN_COL_MONEY)
            If money > 0 Then modState.AddStat modConfig.STAT_MONEY, money

            ' Drop items (pipe-delimited item IDs)
            Dim drops As String
            drops = modData.ReadCellStr(modConfig.SH_ENEMIES, row, EN_COL_DROPS)
            If Len(drops) > 0 Then
                Dim dropItems As Variant
                dropItems = modUtils.SplitTrimmed(drops, "|")
                Dim i As Long
                For i = LBound(dropItems) To UBound(dropItems)
                    Dim dropID As String
                    dropID = Trim(CStr(dropItems(i)))
                    If Len(dropID) > 0 Then
                        modEffects.ProcessEffects "ITEM_ADD:" & dropID
                    End If
                Next i
            End If

            AppendLog "Rewards: +" & xp & " XP, +" & money & " money"
        End If

        ' Reduce rage after victory
        modState.AddStat modConfig.STAT_RAGE, -5
    End If

    If result = RESULT_DEFEAT Then
        ' Player loses some money and rage spikes
        Dim currentMoney As Long
        currentMoney = modState.GetStat(modConfig.STAT_MONEY)
        Dim lostMoney As Long
        lostMoney = currentMoney \ 4
        If lostMoney > 0 Then modState.AddStat modConfig.STAT_MONEY, -CDbl(lostMoney)
        modState.AddStat modConfig.STAT_RAGE, 10
        modState.SetStat modConfig.STAT_HEALTH, 5  ' survive at 5 HP
        AppendLog "You wake up battered... lost " & lostMoney & " money"
    End If

    ' Write final log entry
    WriteCombatLogToSheet result

    modUtils.DebugLog "modCombat.EndCombat: " & mEnemyID & " result=" & result
End Sub

'===============================================================
' PRIVATE — Combat Log Management
'===============================================================

Private Sub AppendLog(entry As String)
    If Len(mCombatLog) > 0 Then
        mCombatLog = mCombatLog & vbLf
    End If
    mCombatLog = mCombatLog & entry
End Sub

' Write a round to the tbl_CombatLog sheet
Private Sub LogCombatRound(actor As String, action As String, dmg As Long, result As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_COMBAT)
    If ws Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = modUtils.GetLastRow(ws, 1) + 1

    ws.Cells(nextRow, 1).Value = "LOG_" & Format(Now, "yyyymmddhhmmss") & "_" & mRound
    ws.Cells(nextRow, 2).Value = Now
    ws.Cells(nextRow, 3).Value = mEnemyID
    ws.Cells(nextRow, 4).Value = mRound
    ws.Cells(nextRow, 5).Value = actor
    ws.Cells(nextRow, 6).Value = action
    ws.Cells(nextRow, 7).Value = dmg
    ws.Cells(nextRow, 8).Value = result
End Sub

' Write a summary entry at the end of combat
Private Sub WriteCombatLogToSheet(finalResult As String)
    Dim ws As Worksheet
    Set ws = modConfig.GetSheet(modConfig.SH_COMBAT)
    If ws Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = modUtils.GetLastRow(ws, 1) + 1

    ws.Cells(nextRow, 1).Value = "LOG_" & Format(Now, "yyyymmddhhmmss") & "_END"
    ws.Cells(nextRow, 2).Value = Now
    ws.Cells(nextRow, 3).Value = mEnemyID
    ws.Cells(nextRow, 4).Value = mRound
    ws.Cells(nextRow, 5).Value = "SYSTEM"
    ws.Cells(nextRow, 6).Value = "COMBAT_END"
    ws.Cells(nextRow, 7).Value = 0
    ws.Cells(nextRow, 8).Value = finalResult
End Sub
