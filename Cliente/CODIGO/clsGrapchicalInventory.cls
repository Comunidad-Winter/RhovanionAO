VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item
Dim base_tile_size As Byte

Dim Slots As Byte

'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory() As Inventory    'User's inventory

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If


Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ base_tile_size)) * (InventoryWindow.ScaleWidth \ base_tile_size) < Slots Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
    
    Render_Inventory = True
    
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property

Public Property Get Def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    Def = UserInventory(slot).Def
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = UserInventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = UserInventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = UserInventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Long, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > Slots Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
    
    Render_Inventory = True
    
End Sub

Private Function ClickItem(ByVal x As Long, ByVal y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    Dim AuxInv As Inventory
    
    temp_x = x \ 32
    temp_y = y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ base_tile_size) + 1
    
    'Make sure it's within limits
    If TempItem <= Slots Then
        'Make sure slot isn't empty
        If UserInventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
    
    Render_Inventory = True
End Function

Public Sub DrawInventory()
    Dim loopc As Byte
    Dim filas As Integer
    Dim x As Long
    Dim y As Long
    Dim h_slots As Byte
    
    h_slots = Fix(InventoryWindow.ScaleWidth / base_tile_size)
    
    For loopc = InventoryOffset * h_slots + 1 To UBound(UserInventory)
        If loopc - filas * h_slots > h_slots Then
            filas = filas + 1
        End If
        
        x = ((loopc - 1) - (filas * h_slots)) * base_tile_size
        y = filas * base_tile_size
        
        If loopc = InvSelectedItem Then
            Call DXEngine_DrawBox(x + 1, y + 1, base_tile_size, base_tile_size, RGB(0, 0, 255))
            Call GUI_Grh_Render(UserInventory(loopc).GrhIndex, x, y, , True, RGB(255, 255, 255))
        Else
            Call GUI_Grh_Render(UserInventory(loopc).GrhIndex, x, y, , True, RGB(150, 150, 150))
        End If
        
        
        
        If UserInventory(loopc).GrhIndex Then
            Call DXEngine_TextRender(2, UserInventory(loopc).Amount, x + 2, y + 1, D3DColorXRGB(255, 255, 255))
            If UserInventory(loopc).Equipped Then
                Call DXEngine_TextRender(2, "+", x + 20, y + 20, D3DColorXRGB(255, 255, 0))
            End If
        End If
        
    Next loopc
End Sub

#If ConMenuesConextuales = 1 Then   'Argument list varies!!!
    Public Sub Initialize(ByRef InvPic As PictureBox, ByRef PopUpMenu As Menu)

#Else   'Argument list varies!!!
    Public Sub Initialize(ByRef InvPic As PictureBox, ByVal NumSlots As Byte)
#End If
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the reference to the picture box on which to render the inventory
'***************************************************
    'Cambiamos la cantidad de items del inventario.
    Slots = NumSlots
    ReDim UserInventory(1 To Slots) As Inventory
    
    Set InventoryWindow = InvPic
    
    base_tile_size = 32
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
    
#If ConMenuesConextuales = 1 Then
    Set ItemMenu = PopUpMenu
#End If
End Sub


Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > InventoryWindow.width Or y > InventoryWindow.height Then _
        Exit Sub
    
    temp_x = x \ base_tile_size
    temp_y = y \ base_tile_size
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ base_tile_size) + 1
    
    If TempItem <= Slots Then
        InventoryWindow.ToolTipText = UserInventory(TempItem).Name
    End If
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    Dim ImagesPath As String
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > InventoryWindow.width Or y > InventoryWindow.height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(x), CInt(y))
            
        If InvSelectedItem Then
            frmMain.lblItemName.Caption = UserInventory(InvSelectedItem).Name
        Else
            frmMain.lblItemName.Caption = ""
        End If
    
#If ConMenuesConextuales = 1 Then
    'If it was a right click on the same item we had selected before, show popup menu
    If Button = vbRightButton And prevSelItem = InvSelectedItem Then
        'Show the provided menu as a popup
        Call InventoryWindow.Parent.PopUpMenu(ItemMenu, , x, y)
    End If
#End If
End Sub

Public Sub ChangePictureBox(NewPictureBox As PictureBox)
'***************************************************
'Author: Nicolás Serafini (ElNiCk)
'Last Modify Date: 1/5/06
'Changes the PictureBox in which we must render.
'***************************************************
    Set InventoryWindow = Nothing
    Set InventoryWindow = NewPictureBox
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    Render_Inventory = True
End Sub

Public Sub Reset()
'***************************************************
'Author: Nicolás Serafini (ElNiCk)
'Last Modify Date: 18/7/06
'Reset Inventory items.
'***************************************************
Dim slot As Integer
For slot = LBound(UserInventory) To UBound(UserInventory)
    Call SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
Next slot
End Sub

Public Sub Render()
    DrawInventory
End Sub
