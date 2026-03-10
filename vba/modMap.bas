Attribute VB_Name = "modMap"
'===============================================================
' modMap — Map Navigation & Travel System
' Damned Moon VBA RPG Engine — Phase 3
'===============================================================
' Handles world-map travel: listing reachable destinations from
' the current node, computing travel time, moving the player,
' and triggering encounters along the way.
'
' Data sources:
'   tbl_MapNodes  — NodeID, Name, Description, Region, Services,
'                   DangerLevel, Requirements
'   tbl_MapLinks  — FromID, ToID, TravelMinutes, DangerMod,
'                   Requirements, Description
'===============================================================

Option Explicit

' ── MAP NODE COLUMN INDICES (tbl_MapNodes) ──
Private Const MN_COL_ID As Long = 1          ' A: NodeID
Private Const MN_COL_NAME As Long = 2        ' B: Display name
Private Const MN_COL_DESC As Long = 3        ' C: Description
Private Const MN_COL_REGION As Long = 4      ' D: Region tag
Private Const MN_COL_SERVICES As Long = 5    ' E: Pipe-delimited services (SHOP|REST|JOB)
Private Const MN_COL_DANGER As Long = 6      ' F: Base danger level (0-100)
Private Const MN_COL_REQS As Long = 7        ' G: Requirements to enter

' ── MAP LINK COLUMN INDICES (tbl_MapLinks) ──
Private Const ML_COL_FROM As Long = 1         ' A: FromID
Private Const ML_COL_TO As Long = 2           ' B: ToID
Private Const ML_COL_MINUTES As Long = 3      ' C: Travel time in minutes
Private Const ML_COL_DANGER_MOD As Long = 4   ' D: Danger modifier for travel
Private Const ML_COL_REQS As Long = 5         ' E: Requirements to use this route
Private Const ML_COL_DESC As Long = 6         ' F: Route description

'===============================================================
' PUBLIC — Get reachable destinations from a node
'===============================================================

' Returns a Collection of destination NodeIDs reachable from
' the given location. Filters by link requirements.
Public Function GetReachableNodes(fromNodeID As String) As Collection
    Dim result As New Collection

    Dim wsLinks As Worksheet
    Set wsLinks = modConfig.GetSheet(modConfig.SH_MAPLINKS)
    If wsLinks Is Nothing Then
        Set GetReachableNodes = result
        Exit Function
    End If

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsLinks, ML_COL_FROM)
        Dim linkFrom As String
        linkFrom = modUtils.SafeStr(wsLinks.Cells(r, ML_COL_FROM).Value)

        If linkFrom = fromNodeID Then
            Dim linkTo As String
            linkTo = modUtils.SafeStr(wsLinks.Cells(r, ML_COL_TO).Value)
            If Len(linkTo) = 0 Then GoTo NextLink

            ' Check link requirements
            Dim linkReqs As String
            linkReqs = modUtils.SafeStr(wsLinks.Cells(r, ML_COL_REQS).Value)
            If Len(linkReqs) > 0 Then
                If Not modRequirements.CheckRequirements(linkReqs) Then GoTo NextLink
            End If

            ' Check destination node requirements
            Dim nodeRow As Long
            nodeRow = modData.GetMapNodeRow(linkTo)
            If nodeRow > 0 Then
                Dim nodeReqs As String
                nodeReqs = modData.ReadCellStr(modConfig.SH_MAPNODES, nodeRow, MN_COL_REQS)
                If Len(nodeReqs) > 0 Then
                    If Not modRequirements.CheckRequirements(nodeReqs) Then GoTo NextLink
                End If
            End If

            result.Add linkTo
        End If
NextLink:
    Next r

    Set GetReachableNodes = result
End Function

'===============================================================
' PUBLIC — Get node display info
'===============================================================

' Get the display name for a map node
Public Function GetNodeName(nodeID As String) As String
    Dim row As Long
    row = modData.GetMapNodeRow(nodeID)
    If row = 0 Then
        GetNodeName = nodeID
        Exit Function
    End If
    GetNodeName = modData.ReadCellStr(modConfig.SH_MAPNODES, row, MN_COL_NAME)
    If Len(GetNodeName) = 0 Then GetNodeName = nodeID
End Function

' Get the description for a map node
Public Function GetNodeDescription(nodeID As String) As String
    Dim row As Long
    row = modData.GetMapNodeRow(nodeID)
    If row = 0 Then Exit Function
    GetNodeDescription = modData.ReadCellStr(modConfig.SH_MAPNODES, row, MN_COL_DESC)
End Function

' Get the base danger level for a node (0-100)
Public Function GetNodeDanger(nodeID As String) As Long
    Dim row As Long
    row = modData.GetMapNodeRow(nodeID)
    If row = 0 Then Exit Function
    GetNodeDanger = modData.ReadCellLng(modConfig.SH_MAPNODES, row, MN_COL_DANGER)
End Function

' Get the region tag for a node
Public Function GetNodeRegion(nodeID As String) As String
    Dim row As Long
    row = modData.GetMapNodeRow(nodeID)
    If row = 0 Then Exit Function
    GetNodeRegion = modData.ReadCellStr(modConfig.SH_MAPNODES, row, MN_COL_REGION)
End Function

' Check if a node offers a specific service (SHOP, REST, JOB, etc.)
Public Function NodeHasService(nodeID As String, serviceName As String) As Boolean
    Dim row As Long
    row = modData.GetMapNodeRow(nodeID)
    If row = 0 Then Exit Function

    Dim services As String
    services = modData.ReadCellStr(modConfig.SH_MAPNODES, row, MN_COL_SERVICES)
    NodeHasService = (InStr(UCase(services), UCase(serviceName)) > 0)
End Function

'===============================================================
' PUBLIC — Get link info between two nodes
'===============================================================

' Get travel time in minutes between two nodes. Returns 0 if no link.
Public Function GetTravelTime(fromID As String, toID As String) As Long
    Dim row As Long
    row = modData.GetMapLinkRow(fromID, toID)
    If row = 0 Then Exit Function
    GetTravelTime = modData.ReadCellLng(modConfig.SH_MAPLINKS, row, ML_COL_MINUTES)
End Function

' Get route description between two nodes
Public Function GetRouteDescription(fromID As String, toID As String) As String
    Dim row As Long
    row = modData.GetMapLinkRow(fromID, toID)
    If row = 0 Then Exit Function
    GetRouteDescription = modData.ReadCellStr(modConfig.SH_MAPLINKS, row, ML_COL_DESC)
End Function

' Get danger modifier for a route
Public Function GetRouteDangerMod(fromID As String, toID As String) As Long
    Dim row As Long
    row = modData.GetMapLinkRow(fromID, toID)
    If row = 0 Then Exit Function
    GetRouteDangerMod = modData.ReadCellLng(modConfig.SH_MAPLINKS, row, ML_COL_DANGER_MOD)
End Function

'===============================================================
' PUBLIC — Travel to a destination
'===============================================================

' Move the player from current location to a destination node.
' Advances time, checks for encounters, updates location & UI.
' Returns True if travel completed, False if blocked or interrupted.
Public Function TravelTo(destNodeID As String) As Boolean
    Dim currentLoc As String
    currentLoc = modState.GetCurrentLocation()

    ' Verify destination is reachable
    If Not IsReachable(currentLoc, destNodeID) Then
        modUtils.DebugLog "modMap.TravelTo: " & destNodeID & " not reachable from " & currentLoc
        TravelTo = False
        Exit Function
    End If

    ' Get travel time
    Dim travelMin As Long
    travelMin = GetTravelTime(currentLoc, destNodeID)
    If travelMin <= 0 Then travelMin = 30  ' default 30 min

    ' Advance time for travel
    modState.AdvanceTime travelMin

    ' Calculate effective danger for the journey
    Dim routeDanger As Long
    routeDanger = GetRouteDangerMod(currentLoc, destNodeID)
    Dim destDanger As Long
    destDanger = GetNodeDanger(destNodeID)
    Dim effectiveDanger As Long
    effectiveDanger = modUtils.Clamp(destDanger + routeDanger, 0, 100)

    ' Night travel is more dangerous
    If modState.IsNight() Then
        effectiveDanger = modUtils.Clamp(effectiveDanger + 20, 0, 100)
    End If

    ' Roll for travel encounter
    Dim encounterID As String
    encounterID = modEncounters.RollEncounter(currentLoc, destNodeID, effectiveDanger)

    ' Update location
    modState.SetCurrentLocation destNodeID

    ' If an encounter was triggered, handle it
    If Len(encounterID) > 0 Then
        modEncounters.RunEncounter encounterID
        modUtils.DebugLog "modMap.TravelTo: encounter " & encounterID & " during travel to " & destNodeID
    End If

    ' Update UI
    modUI.UpdateMapHighlight destNodeID
    modUI.UpdateDayTimePanel
    modUI.UpdateStatsPanel

    modUtils.DebugLog "modMap.TravelTo: arrived at " & destNodeID & " (" & travelMin & " min)"
    TravelTo = True
End Function

'===============================================================
' PUBLIC — Build travel choice text for scene engine
'===============================================================

' Generates a formatted string listing available destinations.
' Used by the world loop to show travel options.
Public Function BuildTravelNarrative(fromNodeID As String) As String
    Dim currentName As String
    currentName = GetNodeName(fromNodeID)
    Dim currentDesc As String
    currentDesc = GetNodeDescription(fromNodeID)

    Dim text As String
    text = "You are at " & currentName & "." & vbLf

    If Len(currentDesc) > 0 Then
        text = text & currentDesc & vbLf
    End If

    text = text & vbLf & "Where will you go?"

    BuildTravelNarrative = text
End Function

' Populate the Game sheet with travel destination choices.
' Returns the number of destinations shown.
Public Function ShowTravelChoices(fromNodeID As String) As Long
    Dim destinations As Collection
    Set destinations = GetReachableNodes(fromNodeID)

    Dim count As Long
    count = 0

    Dim i As Long
    For i = 1 To destinations.Count
        If count >= modConfig.MAX_CHOICES Then Exit For

        Dim destID As String
        destID = CStr(destinations(i))

        Dim destName As String
        destName = GetNodeName(destID)

        Dim travelMin As Long
        travelMin = GetTravelTime(fromNodeID, destID)

        Dim routeDesc As String
        routeDesc = GetRouteDescription(fromNodeID, destID)

        ' Build choice label
        Dim label As String
        label = "Travel to " & destName
        If travelMin > 0 Then
            label = label & "  (" & travelMin & " min)"
        End If
        If Len(routeDesc) > 0 Then
            label = label & "  — " & routeDesc
        End If

        count = count + 1
        modUI.ShowChoiceButton count, CStr(count) & ".  " & label, True
    Next i

    ' Hide remaining buttons
    Dim j As Long
    For j = count + 1 To modConfig.MAX_CHOICES
        modUI.HideChoiceButton j
    Next j

    ShowTravelChoices = count
End Function

'===============================================================
' PUBLIC — Enter the world / free-roam loop
'===============================================================

' Switches the game into map travel mode at the current location.
' Shows the travel narrative and destination choices.
Public Sub EnterWorldMap()
    Application.ScreenUpdating = False

    Dim currentLoc As String
    currentLoc = modState.GetCurrentLocation()

    If Len(currentLoc) = 0 Then
        currentLoc = modConfig.DEFAULT_START_LOCATION
        modState.SetCurrentLocation currentLoc
    End If

    ' Show travel narrative
    Dim narrative As String
    narrative = BuildTravelNarrative(currentLoc)
    modUI.ShowNarrative narrative

    ' Show travel choices
    ShowTravelChoices currentLoc

    ' Update HUD
    modUI.UpdateStatsPanel
    modUI.UpdateQuestPanel
    modUI.UpdateInventoryPanel
    modUI.UpdateDayTimePanel
    modUI.UpdateMapHighlight currentLoc

    modUtils.DebugLog "modMap.EnterWorldMap: showing map at " & currentLoc
    Application.ScreenUpdating = True
End Sub

'===============================================================
' PRIVATE HELPERS
'===============================================================

' Check if a destination is reachable from a source
Private Function IsReachable(fromID As String, toID As String) As Boolean
    Dim destinations As Collection
    Set destinations = GetReachableNodes(fromID)

    Dim i As Long
    For i = 1 To destinations.Count
        If CStr(destinations(i)) = toID Then
            IsReachable = True
            Exit Function
        End If
    Next i

    IsReachable = False
End Function
