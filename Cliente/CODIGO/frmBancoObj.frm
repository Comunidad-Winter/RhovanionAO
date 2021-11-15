VERSION 5.00
Begin VB.Form frmBancoObj 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox invUsu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   3825
      ScaleHeight     =   4005
      ScaleWidth      =   2490
      TabIndex        =   8
      Top             =   1755
      Width           =   2520
   End
   Begin VB.PictureBox InvNpc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   600
      ScaleHeight     =   4005
      ScaleWidth      =   2490
      TabIndex        =   7
      Top             =   1755
      Width           =   2520
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   405
      Left            =   3180
      TabIndex        =   5
      Text            =   "1"
      Top             =   6120
      Width           =   570
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   360
      ScaleHeight     =   690
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   6450
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2670
      TabIndex        =   6
      Top             =   1140
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   1
      Left            =   3780
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6090
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   540
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6090
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MaxGolpe"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3870
      TabIndex        =   4
      Top             =   1050
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MinGolpe"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3870
      TabIndex        =   3
      Top             =   810
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cant"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1860
      TabIndex        =   2
      Top             =   750
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3870
      TabIndex        =   1
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   6510
      TabIndex        =   9
      Top             =   6870
      Width           =   315
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer




Private Sub cantidad_Change()
If Val(cantidad.Text) < 0 Then
    cantidad.Text = 1
End If

If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
    cantidad.Text = 1
End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadPicture(DirGraficos & "comerciar.jpg")
Image1(0).Picture = LoadPicture(DirGraficos & "BotónComprar.jpg")
Image1(1).Picture = LoadPicture(DirGraficos & "Botónvender.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\resources\Graphics\BotónComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\resources\Graphics\Botónvender.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub Image1_Click(Index As Integer)

modSound.Sound_Play SND_CLICK, DSBPLAY_DEFAULT

Select Case Index
    Case 0
    
        If NpcInv.SelectedItem = 0 Then Exit Sub
        'frmBancoObj.List1(0).SetFocus
        'LastIndex1 = List1(0).listIndex
        LasActionBuy = True
        Call WriteBankExtractItem(NpcInv.SelectedItem, cantidad.Text)
        
   Case 1
        
        If Inventario.SelectedItem = 0 Then Exit Sub
        'LastIndex2 = List1(1).listIndex
        LasActionBuy = False
        If Not Inventario.Equipped(Inventario.SelectedItem) Then
            Call WriteBankDeposit(Inventario.SelectedItem, cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes depositar el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select
'List1(0).Clear

'List1(1).Clear

NPCInvDim = 0
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
                Image1(0).Picture = LoadPicture(App.Path & "\resources\Graphics\BotónComprarApretado.jpg")
                Image1(0).Tag = 0
                Image1(1).Picture = LoadPicture(App.Path & "\resources\Graphics\Botónvender.jpg")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = LoadPicture(App.Path & "\resources\Graphics\Botónvenderapretado.jpg")
                Image1(1).Tag = 0
                Image1(0).Picture = LoadPicture(App.Path & "\resources\Graphics\BotónComprar.jpg")
                Image1(0).Tag = 1
        End If
        
End Select
End Sub

Private Sub Image2_Click()
    Call WriteBankEnd
End Sub

Private Sub InvNpc_Click()

If NpcInv.SelectedItem = 0 Then Exit Sub

Label1(0).Caption = UserBancoInventory(NpcInv.SelectedItem).Name
        Label1(2).Caption = "Cantidad: " & UserBancoInventory(NpcInv.SelectedItem).Amount
        Select Case UserBancoInventory(NpcInv.SelectedItem).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & UserBancoInventory(NpcInv.SelectedItem).MaxHit
                Label1(4).Caption = "Min Golpe:" & UserBancoInventory(NpcInv.SelectedItem).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & UserBancoInventory(NpcInv.SelectedItem).Def
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        'Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, UserBancoInventory(List1(0).listIndex + 1).GrhIndex, SR, DR)
End Sub

Private Sub invUsu_Click()

If Inventario.SelectedItem = 0 Then Exit Sub

Label1(0).Caption = Inventario.ItemName(Inventario.SelectedItem)
        Label1(2).Caption = "Cantidad: " & Inventario.Amount(Inventario.SelectedItem)
        Select Case Inventario.OBJType(Inventario.SelectedItem)
            Case 2
                Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(Inventario.SelectedItem)
                Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(Inventario.SelectedItem)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.Def(Inventario.SelectedItem)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
End Sub
