Attribute VB_Name = "modInventory"
'===============================================================
' modInventory — Inventory Management System
' Damned Moon VBA RPG Engine — Phase 4
'===============================================================
' Full inventory CRUD: add, remove, equip, unequip, use items.
' Items have types (WEAPON, CHARM, COAT, CONSUMABLE, KEY, JUNK),
' equip slots, stack limits, passive effects, and use effects.
'
' Data sources:
'   tbl_ItemDB    — item definitions (stats, types, effects)
'   tbl_Inventory — player's current inventory slots
'
' tbl_ItemDB columns:
'   A: ItemID       B: Name       C: Description
'   D: Type         E: Stackable  F: MaxStack
'   G: EquipSlot    H: PassiveEffect   I: UseEffect
'   J: Value        K: Weight     L: Rarity
'
' tbl_Inventory columns:
'   A: SlotNum  B: ItemID  C: ItemName  D: Qty  E: Equipped
'===============================================================

Option Explicit

' ── ITEM DB COLUMN INDICES ──
Private Const IDB_COL_ID As Long = 1           ' A: ItemID
Private Const IDB_COL_NAME As Long = 2         ' B: Display name
Private Const IDB_COL_DESC As Long = 3         ' C: Description
Private Const IDB_COL_TYPE As Long = 4         ' D: Type (WEAPON, CHARM, COAT, CONSUMABLE, KEY, JUNK)
Private Const IDB_COL_STACKABLE As Long = 5    ' E: Stackable (TRUE/FALSE)
Private Const IDB_COL_MAXSTACK As Long = 6     ' F: Max stack size
Private Const IDB_COL_EQUIPSLOT As Long = 7    ' G: Equip slot (WEAPON, CHARM, COAT, or empty)
Private Const IDB_COL_PASSIVE As Long = 8      ' H: Passive effect string (applied while equipped)
Private Const IDB_COL_USE As Long = 9          ' I: Use/consumable effect string
Private Const IDB_COL_VALUE As Long = 10       ' J: Monetary value
Private Const IDB_COL_WEIGHT As Long = 11      ' K: Weight (for carry limit)
Private Const IDB_COL_RARITY As Long = 12      ' L: Rarity (COMMON, UNCOMMON, RARE, UNIQUE)

' ── INVENTORY COLUMN INDICES ──
Private Const INV_COL_SLOT As Long = 1         ' A: Slot number
Private Const INV_COL_ITEMID As Long = 2       ' B: ItemID
Private Const INV_COL_NAME As Long = 3         ' C: Item display name
Private Const INV_COL_QTY As Long = 4          ' D: Quantity
Private Const INV_COL_EQUIPPED As Long = 5     ' E: Equipped flag

' ── EQUIP SLOT NAMES ──
Public Const SLOT_WEAPON As String = "WEAPON"
Public Const SLOT_CHARM As String = "CHARM"
Public Const SLOT_COAT As String = "COAT"

'===============================================================
' PUBLIC — Add item to inventory
'===============================================================

' Add an item by ID. Returns True if successfully added.
Public Function AddItem(itemID As String, Optional qty As Long = 1) As Boolean
    AddItem = False
    If Len(itemID) = 0 Or qty <= 0 Then Exit Function

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Function

    ' Look up item definition
    Dim itemName As String
    itemName = GetItemName(itemID)
    If Len(itemName) = 0 Then itemName = itemID

    ' Check if already in inventory (try to stack)
    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, INV_COL_SLOT)
        If modUtils.SafeStr(wsInv.Cells(r, INV_COL_ITEMID).Value) = itemID Then
            ' Item found — check stacking
            If IsStackable(itemID) Then
                Dim currentQty As Long
                currentQty = modUtils.SafeLng(wsInv.Cells(r, INV_COL_QTY).Value, 0)
                Dim maxStack As Long
                maxStack = GetMaxStack(itemID)
                Dim addQty As Long
                addQty = qty
                If maxStack > 0 And (currentQty + addQty) > maxStack Then
                    addQty = maxStack - currentQty
                End If
                If addQty > 0 Then
                    wsInv.Cells(r, INV_COL_QTY).Value = currentQty + addQty
                    AddItem = True
                    modUtils.DebugLog "modInventory.AddItem: stacked " & itemID & " +" & addQty
                End If
            Else
                modUtils.DebugLog "modInventory.AddItem: " & itemID & " not stackable, already owned"
            End If
            Exit Function
        End If
    Next r

    ' Not in inventory — find empty slot
    For r = 2 To modUtils.GetLastRow(wsInv, INV_COL_SLOT)
        Dim slotID As String
        slotID = modUtils.SafeStr(wsInv.Cells(r, INV_COL_ITEMID).Value)
        If Len(slotID) = 0 Then
            wsInv.Cells(r, INV_COL_ITEMID).Value = itemID
            wsInv.Cells(r, INV_COL_NAME).Value = itemName
            wsInv.Cells(r, INV_COL_QTY).Value = qty
            wsInv.Cells(r, INV_COL_EQUIPPED).Value = False
            AddItem = True
            modUtils.DebugLog "modInventory.AddItem: added " & itemID & " x" & qty
            Exit Function
        End If
    Next r

    modUtils.DebugLog "modInventory.AddItem: no empty slot for " & itemID
End Function

'===============================================================
' PUBLIC — Remove item from inventory
'===============================================================

' Remove an item by ID. Returns True if successfully removed.
Public Function RemoveItem(itemID As String, Optional qty As Long = 1) As Boolean
    RemoveItem = False
    If Len(itemID) = 0 Or qty <= 0 Then Exit Function

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Function

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, INV_COL_SLOT)
        If modUtils.SafeStr(wsInv.Cells(r, INV_COL_ITEMID).Value) = itemID Then
            Dim currentQty As Long
            currentQty = modUtils.SafeLng(wsInv.Cells(r, INV_COL_QTY).Value, 0)

            ' Unequip if equipped and removing all
            If currentQty <= qty Then
                If modUtils.SafeBool(wsInv.Cells(r, INV_COL_EQUIPPED).Value) Then
                    UnequipFromSlot r, wsInv
                End If
                ' Clear the slot entirely
                wsInv.Cells(r, INV_COL_ITEMID).Value = ""
                wsInv.Cells(r, INV_COL_NAME).Value = ""
                wsInv.Cells(r, INV_COL_QTY).Value = 0
                wsInv.Cells(r, INV_COL_EQUIPPED).Value = False
            Else
                wsInv.Cells(r, INV_COL_QTY).Value = currentQty - qty
            End If

            RemoveItem = True
            modUtils.DebugLog "modInventory.RemoveItem: " & itemID & " -" & qty
            Exit Function
        End If
    Next r

    modUtils.DebugLog "modInventory.RemoveItem: " & itemID & " not found in inventory"
End Function

'===============================================================
' PUBLIC — Has item check
'===============================================================

' Check if the player has at least N of an item
Public Function HasItem(itemID As String, Optional minQty As Long = 1) As Boolean
    HasItem = (GetItemQty(itemID) >= minQty)
End Function

' Get the quantity of an item in inventory
Public Function GetItemQty(itemID As String) As Long
    GetItemQty = 0

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Function

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, INV_COL_SLOT)
        If modUtils.SafeStr(wsInv.Cells(r, INV_COL_ITEMID).Value) = itemID Then
            GetItemQty = modUtils.SafeLng(wsInv.Cells(r, INV_COL_QTY).Value, 0)
            Exit Function
        End If
    Next r
End Function

'===============================================================
' PUBLIC — Equip / Unequip
'===============================================================

' Equip an item from inventory. Returns True on success.
Public Function EquipItem(itemID As String) As Boolean
    EquipItem = False
    If Len(itemID) = 0 Then Exit Function

    ' Check item has an equip slot
    Dim equipSlot As String
    equipSlot = GetItemEquipSlot(itemID)
    If Len(equipSlot) = 0 Then
        modUtils.DebugLog "modInventory.EquipItem: " & itemID & " has no equip slot"
        Exit Function
    End If

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Function

    ' Find item in inventory
    Dim itemRow As Long
    itemRow = FindInventoryRow(itemID, wsInv)
    If itemRow = 0 Then
        modUtils.DebugLog "modInventory.EquipItem: " & itemID & " not in inventory"
        Exit Function
    End If

    ' Unequip current item in same slot (if any)
    UnequipSlot equipSlot

    ' Equip the item
    wsInv.Cells(itemRow, INV_COL_EQUIPPED).Value = True

    ' Apply passive effects
    Dim passiveEff As String
    passiveEff = GetItemPassiveEffect(itemID)
    If Len(passiveEff) > 0 Then
        modEffects.ProcessEffects passiveEff
    End If

    EquipItem = True
    modUtils.DebugLog "modInventory.EquipItem: equipped " & itemID & " in " & equipSlot
End Function

' Unequip an item. Returns True on success.
Public Function UnequipItem(itemID As String) As Boolean
    UnequipItem = False

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Function

    Dim itemRow As Long
    itemRow = FindInventoryRow(itemID, wsInv)
    If itemRow = 0 Then Exit Function

    If Not modUtils.SafeBool(wsInv.Cells(itemRow, INV_COL_EQUIPPED).Value) Then
        Exit Function  ' Not equipped
    End If

    UnequipFromSlot itemRow, wsInv
    UnequipItem = True
End Function

' Unequip whatever is in a given equip slot
Public Sub UnequipSlot(slotName As String)
    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Sub

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, INV_COL_SLOT)
        If modUtils.SafeBool(wsInv.Cells(r, INV_COL_EQUIPPED).Value) Then
            Dim eqID As String
            eqID = modUtils.SafeStr(wsInv.Cells(r, INV_COL_ITEMID).Value)
            If UCase(GetItemEquipSlot(eqID)) = UCase(slotName) Then
                UnequipFromSlot r, wsInv
                Exit Sub
            End If
        End If
    Next r
End Sub

'===============================================================
' PUBLIC — Use a consumable item
'===============================================================

' Use a consumable item. Applies its UseEffect and decrements qty.
' Returns True if used successfully.
Public Function UseItem(itemID As String) As Boolean
    UseItem = False
    If Len(itemID) = 0 Then Exit Function

    ' Must be a consumable
    If UCase(GetItemType(itemID)) <> "CONSUMABLE" Then
        modUtils.DebugLog "modInventory.UseItem: " & itemID & " is not consumable"
        Exit Function
    End If

    ' Must have it
    If Not HasItem(itemID) Then
        modUtils.DebugLog "modInventory.UseItem: " & itemID & " not in inventory"
        Exit Function
    End If

    ' Get use effect
    Dim useEff As String
    useEff = GetItemUseEffect(itemID)

    ' Apply effect
    If Len(useEff) > 0 Then
        modEffects.ProcessEffects useEff
    End If

    ' Consume (remove 1)
    RemoveItem itemID, 1

    UseItem = True
    modUtils.DebugLog "modInventory.UseItem: used " & itemID
End Function

'===============================================================
' PUBLIC — Item DB lookups
'===============================================================

' Get item display name from DB
Public Function GetItemName(itemID As String) As String
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    GetItemName = modData.ReadCellStr(modConfig.SH_ITEMS, row, IDB_COL_NAME)
End Function

' Get item description from DB
Public Function GetItemDescription(itemID As String) As String
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    GetItemDescription = modData.ReadCellStr(modConfig.SH_ITEMS, row, IDB_COL_DESC)
End Function

' Get item type from DB
Public Function GetItemType(itemID As String) As String
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    GetItemType = UCase(modData.ReadCellStr(modConfig.SH_ITEMS, row, IDB_COL_TYPE))
End Function

' Get item equip slot from DB (WEAPON, CHARM, COAT, or "")
Public Function GetItemEquipSlot(itemID As String) As String
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    GetItemEquipSlot = UCase(modData.ReadCellStr(modConfig.SH_ITEMS, row, IDB_COL_EQUIPSLOT))
End Function

' Get item passive effect string (active while equipped)
Public Function GetItemPassiveEffect(itemID As String) As String
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    GetItemPassiveEffect = modData.ReadCellStr(modConfig.SH_ITEMS, row, IDB_COL_PASSIVE)
End Function

' Get item use/consumable effect string
Public Function GetItemUseEffect(itemID As String) As String
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    GetItemUseEffect = modData.ReadCellStr(modConfig.SH_ITEMS, row, IDB_COL_USE)
End Function

' Get item monetary value
Public Function GetItemValue(itemID As String) As Long
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    GetItemValue = modData.ReadCellLng(modConfig.SH_ITEMS, row, IDB_COL_VALUE)
End Function

' Get item rarity
Public Function GetItemRarity(itemID As String) As String
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    GetItemRarity = UCase(modData.ReadCellStr(modConfig.SH_ITEMS, row, IDB_COL_RARITY))
End Function

' Check if an item is stackable
Public Function IsStackable(itemID As String) As Boolean
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then Exit Function
    IsStackable = modUtils.SafeBool(modData.ReadCell(modConfig.SH_ITEMS, row, IDB_COL_STACKABLE))
End Function

' Get max stack size for an item
Public Function GetMaxStack(itemID As String) As Long
    Dim row As Long
    row = modData.GetItemRow(itemID)
    If row = 0 Then
        GetMaxStack = 1
        Exit Function
    End If
    GetMaxStack = modData.ReadCellLng(modConfig.SH_ITEMS, row, IDB_COL_MAXSTACK)
    If GetMaxStack <= 0 Then GetMaxStack = 99
End Function

'===============================================================
' PUBLIC — Equipped item queries
'===============================================================

' Get the ItemID equipped in a given slot (or "" if empty)
Public Function GetEquippedInSlot(slotName As String) As String
    GetEquippedInSlot = ""

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Function

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, INV_COL_SLOT)
        If modUtils.SafeBool(wsInv.Cells(r, INV_COL_EQUIPPED).Value) Then
            Dim eqID As String
            eqID = modUtils.SafeStr(wsInv.Cells(r, INV_COL_ITEMID).Value)
            If UCase(GetItemEquipSlot(eqID)) = UCase(slotName) Then
                GetEquippedInSlot = eqID
                Exit Function
            End If
        End If
    Next r
End Function

' Check if a specific item is currently equipped
Public Function IsEquipped(itemID As String) As Boolean
    IsEquipped = False

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then Exit Function

    Dim itemRow As Long
    itemRow = FindInventoryRow(itemID, wsInv)
    If itemRow = 0 Then Exit Function
    IsEquipped = modUtils.SafeBool(wsInv.Cells(itemRow, INV_COL_EQUIPPED).Value)
End Function

'===============================================================
' PUBLIC — Inventory listing
'===============================================================

' Get a Collection of all item IDs currently in inventory
Public Function GetAllItems() As Collection
    Dim result As New Collection

    Dim wsInv As Worksheet
    Set wsInv = modConfig.GetSheet(modConfig.SH_INV)
    If wsInv Is Nothing Then
        Set GetAllItems = result
        Exit Function
    End If

    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, INV_COL_SLOT)
        Dim iid As String
        iid = modUtils.SafeStr(wsInv.Cells(r, INV_COL_ITEMID).Value)
        If Len(iid) > 0 Then
            result.Add iid
        End If
    Next r

    Set GetAllItems = result
End Function

' Count total distinct items in inventory
Public Function GetItemCount() As Long
    GetItemCount = GetAllItems().Count
End Function

'===============================================================
' PRIVATE — Helpers
'===============================================================

' Find the inventory row for an item ID. Returns 0 if not found.
Private Function FindInventoryRow(itemID As String, wsInv As Worksheet) As Long
    FindInventoryRow = 0
    Dim r As Long
    For r = 2 To modUtils.GetLastRow(wsInv, INV_COL_SLOT)
        If modUtils.SafeStr(wsInv.Cells(r, INV_COL_ITEMID).Value) = itemID Then
            FindInventoryRow = r
            Exit Function
        End If
    Next r
End Function

' Unequip an item at a specific inventory row (clears equipped flag, reverses passive)
Private Sub UnequipFromSlot(invRow As Long, wsInv As Worksheet)
    Dim eqID As String
    eqID = modUtils.SafeStr(wsInv.Cells(invRow, INV_COL_ITEMID).Value)

    wsInv.Cells(invRow, INV_COL_EQUIPPED).Value = False
    modUtils.DebugLog "modInventory.UnequipFromSlot: unequipped " & eqID
End Sub
