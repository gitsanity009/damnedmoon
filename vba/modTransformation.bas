Attribute VB_Name = "modTransformation"
'===============================================================
' modTransformation — Werewolf Transformation System
' Damned Moon VBA RPG Engine — Phase 4
'===============================================================
' Multi-stage lycanthropic transformation with control checks.
' Transformation progresses through stages driven by Rage, Hunger,
' moon phase, and time of day. Control (Humanity - Rage) determines
' whether Marcus keeps agency or blacks out.
'
' Transformation stages:
'   0 = HUMAN     — No symptoms
'   1 = ITCH      — Skin crawling, heightened senses
'   2 = CRACK     — Bones shifting, partial features
'   3 = SURGE     — Near-full, violent urges, last chance
'   4 = BLACKOUT  — Full transformation, lose control
'   5 = AFTERMATH — Wake up, assess damage
'
' Triggers:
'   - Rage >= thresholds
'   - Full moon + night
'   - Hunger >= 80 at night
'   - Manual via effect: TRANSFORM_CHECK
'
' Control checks:
'   Control = Humanity - Rage (clamped 0-100)
'   Roll d100 <= Control to resist advancing a stage
'   Modifiers: moon phase, hunger, time of day, items
'===============================================================

Option Explicit

' ── TRANSFORMATION STAGES ──
Public Const STAGE_HUMAN As Long = 0
Public Const STAGE_ITCH As Long = 1
Public Const STAGE_CRACK As Long = 2
Public Const STAGE_SURGE As Long = 3
Public Const STAGE_BLACKOUT As Long = 4
Public Const STAGE_AFTERMATH As Long = 5

' ── STAT/FLAG NAMES ──
Private Const STAT_TRANSFORM_STAGE As String = "TRANSFORM_STAGE"
Private Const FLAG_TRANSFORMED_TODAY As String = "TransformedToday"
Private Const FLAG_BLACKOUT_ACTIVE As String = "BlackoutActive"

' ── RAGE THRESHOLDS ──
Private Const RAGE_ITCH_THRESHOLD As Long = 40
Private Const RAGE_CRACK_THRESHOLD As Long = 60
Private Const RAGE_SURGE_THRESHOLD As Long = 80
Private Const RAGE_BLACKOUT_THRESHOLD As Long = 100

' ── CONTROL CHECK DIFFICULTY ──
Private Const BASE_CONTROL_DC As Long = 50  ' base difficulty

' ── BLACKOUT OUTCOME TABLE ──
' Weighted outcomes for what happens during blackout
Private Const OUTCOME_COUNT As Long = 6

'===============================================================
' PUBLIC — Check transformation pressure
'===============================================================

' Called each scene load / time advance to check if transformation
' should progress. Returns True if stage changed.
Public Function CheckTransformation() As Boolean
    CheckTransformation = False

    Dim currentStage As Long
    currentStage = GetCurrentStage()

    ' Already in blackout or aftermath — don't re-trigger
    If currentStage >= STAGE_BLACKOUT Then Exit Function

    Dim rage As Long
    rage = modState.GetStat(modConfig.STAT_RAGE)
    Dim hunger As Long
    hunger = modState.GetStat(modConfig.STAT_HUNGER)
    Dim isNight As Boolean
    isNight = modState.IsNight()
    Dim isFullMoon As Boolean
    isFullMoon = modState.IsFullMoon()

    ' Determine target stage based on current conditions
    Dim targetStage As Long
    targetStage = STAGE_HUMAN

    If rage >= RAGE_BLACKOUT_THRESHOLD Then
        targetStage = STAGE_BLACKOUT
    ElseIf rage >= RAGE_SURGE_THRESHOLD Then
        targetStage = STAGE_SURGE
    ElseIf rage >= RAGE_CRACK_THRESHOLD Then
        targetStage = STAGE_CRACK
    ElseIf rage >= RAGE_ITCH_THRESHOLD Then
        targetStage = STAGE_ITCH
    End If

    ' Full moon at night forces minimum CRACK stage
    If isFullMoon And isNight Then
        If targetStage < STAGE_CRACK Then targetStage = STAGE_CRACK
    End If

    ' High hunger at night escalates
    If hunger >= 80 And isNight Then
        If targetStage < STAGE_ITCH Then targetStage = STAGE_ITCH
        ' Hunger adds pressure to advance further
        If targetStage < STAGE_SURGE And hunger >= 90 Then
            targetStage = targetStage + 1
        End If
    End If

    ' Only advance (never retreat automatically)
    If targetStage <= currentStage Then Exit Function

    ' Attempt control check to resist
    If targetStage < STAGE_BLACKOUT Then
        If RollControlCheck(targetStage) Then
            ' Resisted — stay at current stage
            modUtils.DebugLog "modTransformation.Check: resisted advance to stage " & targetStage
            Exit Function
        End If
    End If

    ' Advance to target stage
    AdvanceToStage targetStage
    CheckTransformation = True
End Function

'===============================================================
' PUBLIC — Get / Set stage
'===============================================================

' Get the current transformation stage (0-5)
Public Function GetCurrentStage() As Long
    GetCurrentStage = modState.GetStat(STAT_TRANSFORM_STAGE)
    If GetCurrentStage < STAGE_HUMAN Then GetCurrentStage = STAGE_HUMAN
    If GetCurrentStage > STAGE_AFTERMATH Then GetCurrentStage = STAGE_AFTERMATH
End Function

' Get the display name for a stage
Public Function GetStageName(stage As Long) As String
    Select Case stage
        Case STAGE_HUMAN: GetStageName = "Human"
        Case STAGE_ITCH: GetStageName = "The Itch"
        Case STAGE_CRACK: GetStageName = "The Crack"
        Case STAGE_SURGE: GetStageName = "The Surge"
        Case STAGE_BLACKOUT: GetStageName = "Blackout"
        Case STAGE_AFTERMATH: GetStageName = "Aftermath"
        Case Else: GetStageName = "Unknown"
    End Select
End Function

' Reset transformation to human (called on day advance / rest)
Public Sub ResetTransformation()
    modState.SetStat STAT_TRANSFORM_STAGE, STAGE_HUMAN
    modState.SetFlag FLAG_BLACKOUT_ACTIVE, False
    modUtils.DebugLog "modTransformation.ResetTransformation: returned to human"
End Sub

'===============================================================
' PUBLIC — Control check
'===============================================================

' Roll a control check against a target stage.
' Returns True if the player RESISTS the transformation.
Public Function RollControlCheck(targetStage As Long) As Boolean
    Dim control As Long
    control = modState.GetControl()

    ' Calculate difficulty
    Dim difficulty As Long
    difficulty = BASE_CONTROL_DC

    ' Stage pressure: higher stages are harder to resist
    difficulty = difficulty + (targetStage * 10)

    ' Moon modifier: full moon adds +20 difficulty
    If modState.IsFullMoon() Then difficulty = difficulty + 20

    ' Night modifier: +10 difficulty at night
    If modState.IsNight() Then difficulty = difficulty + 10

    ' Hunger modifier: high hunger makes it harder
    Dim hunger As Long
    hunger = modState.GetStat(modConfig.STAT_HUNGER)
    If hunger >= 80 Then difficulty = difficulty + 15
    If hunger >= 60 Then difficulty = difficulty + 5

    ' Composure bonus: high composure helps resist
    Dim composure As Long
    composure = modState.GetStat(modConfig.STAT_COMPOSURE)
    control = control + (composure \ 4)

    ' Silver item bonus: equipped silver items help
    If modInventory.HasItem("ITM_SILVER_KNIFE") Then control = control + 10
    If Len(modInventory.GetEquippedInSlot(modInventory.SLOT_CHARM)) > 0 Then
        control = control + 5
    End If

    ' Suppressant effect
    If modState.GetFlag("UsedSuppressant") Then
        control = control + 25
        modState.SetFlag "UsedSuppressant", False
    End If

    ' Roll d100
    Dim roll As Long
    roll = modUtils.RandBetween(1, 100)

    ' Success if roll <= (control - difficulty + 50), clamped to 5-95%
    Dim successChance As Long
    successChance = modUtils.Clamp(control - difficulty + 50, 5, 95)

    RollControlCheck = (roll <= successChance)

    modUtils.DebugLog "modTransformation.RollControlCheck: roll=" & roll & _
        " vs " & successChance & "% (control=" & control & ", diff=" & difficulty & ")" & _
        IIf(RollControlCheck, " RESISTED", " FAILED")
End Function

'===============================================================
' PUBLIC — Force transformation (from effects)
'===============================================================

' Force advance to a specific stage (used by TRANSFORM effect)
Public Sub ForceStage(stage As Long)
    If stage < STAGE_HUMAN Then stage = STAGE_HUMAN
    If stage > STAGE_BLACKOUT Then stage = STAGE_BLACKOUT
    AdvanceToStage stage
End Sub

' Trigger a full blackout immediately
Public Sub TriggerBlackout()
    AdvanceToStage STAGE_BLACKOUT
End Sub

'===============================================================
' PUBLIC — Aftermath resolution
'===============================================================

' Resolve a blackout: determine what happened while transformed.
' Returns a narrative text describing the aftermath.
Public Function ResolveBlackout() As String
    Dim outcome As Long
    outcome = RollBlackoutOutcome()

    Dim narrativeText As String
    Dim effectStr As String

    Select Case outcome
        Case 1  ' Mild: woke up in the woods
            narrativeText = "You wake in the woods, naked and shivering. " & _
                "Mud cakes your skin. No blood. You got lucky this time."
            effectStr = "STAT:HEALTH-5|STAT:RAGE-30|STAT:HUNGER-20"

        Case 2  ' Animal kill: fed on wildlife
            narrativeText = "The taste of raw meat lingers. Feathers and fur " & _
                "cling to your hands. At least it wasn't human."
            effectStr = "STAT:HUNGER-40|STAT:RAGE-25|STAT:HUMANITY-5|STAT:INSTINCT+10"

        Case 3  ' Violent: destruction of property
            narrativeText = "A barn wall is shattered. Claw marks gouge the " & _
                "wood. Someone will notice. Someone will ask questions."
            effectStr = "STAT:RAGE-20|STAT:HUNGER-10|FLAG_SET:PropertyDestroyed"

        Case 4  ' Dangerous: confronted a human
            narrativeText = "A face. Screaming. Running. You remember the " & _
                "terror in their eyes. Did they see what you are?"
            effectStr = "STAT:RAGE-15|STAT:HUMANITY-10|STAT:COMPOSURE-10"

        Case 5  ' Severe: injured someone
            narrativeText = "Blood on your hands. Not yours. The scent is " & _
                "unmistakable — human. You don't know who. You don't " & _
                "want to know."
            effectStr = "STAT:HEALTH-10|STAT:HUMANITY-20|STAT:RAGE-30|STAT:HUNGER-30|FLAG_SET:InjuredSomeone"

        Case 6  ' Worst: killed
            narrativeText = "The body lies still in the underbrush. Cold. " & _
                "You remember the chase. The catch. The silence after. " & _
                "There is no coming back from this."
            effectStr = "STAT:HUMANITY-35|STAT:RAGE-40|STAT:HUNGER-50|STAT:COMPOSURE-20|FLAG_SET:KilledDuringBlackout"
    End Select

    ' Apply effects
    If Len(effectStr) > 0 Then
        modEffects.ProcessEffects effectStr
    End If

    ' Move to aftermath stage
    modState.SetStat STAT_TRANSFORM_STAGE, STAGE_AFTERMATH
    modState.SetFlag FLAG_BLACKOUT_ACTIVE, False
    modState.SetFlag FLAG_TRANSFORMED_TODAY, True

    modUtils.DebugLog "modTransformation.ResolveBlackout: outcome " & outcome

    ResolveBlackout = narrativeText
End Function

'===============================================================
' PUBLIC — Stage narrative text
'===============================================================

' Get the narrative description for a transformation stage transition
Public Function GetStageNarrative(stage As Long) As String
    Select Case stage
        Case STAGE_ITCH
            GetStageNarrative = "Your skin prickles. Every hair stands on end. " & _
                "The familiar itch crawls up your spine — the wolf is stirring."

        Case STAGE_CRACK
            GetStageNarrative = "A sound like snapping twigs echoes inside your " & _
                "chest. Your jaw aches. Your fingers twitch and curl against " & _
                "your will. The change is coming."

        Case STAGE_SURGE
            GetStageNarrative = "Your vision sharpens to razor clarity. Every " & _
                "heartbeat thunders. The wolf surges against the walls of " & _
                "your mind, clawing for release. Fight it. FIGHT IT."

        Case STAGE_BLACKOUT
            GetStageNarrative = "The world dissolves into red and black. Your " & _
                "last human thought scatters like leaves in a gale. The " & _
                "wolf takes everything." & vbLf & vbLf & _
                "..." & vbLf & vbLf & _
                "Darkness."

        Case STAGE_AFTERMATH
            GetStageNarrative = "Consciousness returns in fragments. Cold air. " & _
                "Aching muscles. The copper taste in your mouth."

        Case Else
            GetStageNarrative = ""
    End Select
End Function

'===============================================================
' PUBLIC — Transformation status queries
'===============================================================

' Check if player is currently transformed (stage >= BLACKOUT)
Public Function IsTransformed() As Boolean
    IsTransformed = (GetCurrentStage() >= STAGE_BLACKOUT)
End Function

' Check if player is in any non-human stage
Public Function IsChanging() As Boolean
    IsChanging = (GetCurrentStage() > STAGE_HUMAN And GetCurrentStage() < STAGE_BLACKOUT)
End Function

' Get the current transformation pressure as a percentage (0-100)
Public Function GetTransformPressure() As Long
    Dim rage As Long
    rage = modState.GetStat(modConfig.STAT_RAGE)
    Dim hunger As Long
    hunger = modState.GetStat(modConfig.STAT_HUNGER)
    Dim moonBonus As Long
    If modState.IsFullMoon() Then moonBonus = 20
    Dim nightBonus As Long
    If modState.IsNight() Then nightBonus = 10

    GetTransformPressure = modUtils.Clamp( _
        (rage * 6 + hunger * 3) \ 10 + moonBonus + nightBonus, 0, 100)
End Function

'===============================================================
' PRIVATE — Stage advancement
'===============================================================

' Advance to a target stage, showing narrative for each step
Private Sub AdvanceToStage(targetStage As Long)
    Dim currentStage As Long
    currentStage = GetCurrentStage()

    If targetStage <= currentStage Then Exit Sub

    ' Step through each intermediate stage for narrative
    Dim stage As Long
    For stage = currentStage + 1 To targetStage
        modState.SetStat STAT_TRANSFORM_STAGE, stage

        ' Show stage narrative
        Dim narrative As String
        narrative = GetStageNarrative(stage)
        If Len(narrative) > 0 Then
            modUI.ShowNarrative narrative
        End If

        ' Stage-specific effects
        ApplyStageEffects stage
    Next stage

    modUtils.DebugLog "modTransformation.AdvanceToStage: " & currentStage & " -> " & targetStage
End Sub

' Apply automatic effects when entering a stage
Private Sub ApplyStageEffects(stage As Long)
    Select Case stage
        Case STAGE_ITCH
            ' Heightened senses
            modState.AddStat modConfig.STAT_INSTINCT, 5

        Case STAGE_CRACK
            ' Partial shift — stat changes
            modState.AddStat modConfig.STAT_INSTINCT, 10
            modState.AddStat modConfig.STAT_COMPOSURE, -10

        Case STAGE_SURGE
            ' Near-full — severe stat pressure
            modState.AddStat modConfig.STAT_INSTINCT, 15
            modState.AddStat modConfig.STAT_COMPOSURE, -15
            modState.AddStat modConfig.STAT_HEALTH, -5

        Case STAGE_BLACKOUT
            ' Full transformation
            modState.SetFlag FLAG_BLACKOUT_ACTIVE, True
            modState.AddStat modConfig.STAT_INSTINCT, 20
    End Select
End Sub

'===============================================================
' PRIVATE — Blackout outcome roller
'===============================================================

' Roll a weighted random blackout outcome (1-6).
' Outcome severity is influenced by Humanity and Control.
Private Function RollBlackoutOutcome() As Long
    Dim humanity As Long
    humanity = modState.GetStat(modConfig.STAT_HUMANITY)
    Dim control As Long
    control = modState.GetControl()

    ' Build weights: higher humanity/control = more weight on mild outcomes
    Dim weights As New Collection

    ' Outcome 1 (mild): base 30, +humanity/2
    weights.Add modUtils.Clamp(30 + humanity \ 2, 5, 80)
    ' Outcome 2 (animal): base 25, +control/3
    weights.Add modUtils.Clamp(25 + control \ 3, 5, 60)
    ' Outcome 3 (property): base 20
    weights.Add CLng(20)
    ' Outcome 4 (confronted): base 15, higher if low humanity
    weights.Add modUtils.Clamp(15 + (100 - humanity) \ 5, 5, 40)
    ' Outcome 5 (injured): base 8, higher if very low control
    weights.Add modUtils.Clamp(8 + (100 - control) \ 4, 2, 30)
    ' Outcome 6 (killed): base 2, only if humanity very low
    Dim killWeight As Long
    If humanity < 20 Then
        killWeight = modUtils.Clamp(10 + (20 - humanity), 2, 25)
    Else
        killWeight = 2
    End If
    weights.Add killWeight

    RollBlackoutOutcome = modUtils.WeightedPick(weights)
    If RollBlackoutOutcome < 1 Or RollBlackoutOutcome > OUTCOME_COUNT Then
        RollBlackoutOutcome = 1  ' fallback to mild
    End If
End Function
