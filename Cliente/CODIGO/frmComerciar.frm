VERSION 5.00
Begin VB.Form frmComerciar 
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
   Begin VB.PictureBox NpcPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4035
      Left            =   600
      ScaleHeight     =   4035
      ScaleWidth      =   2520
      TabIndex        =   7
      Top             =   1755
      Width           =   2520
   End
   Begin VB.PictureBox UsuInv 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4035
      Left            =   3840
      ScaleHeight     =   4035
      ScaleWidth      =   2520
      TabIndex        =   6
      Top             =   1755
      Width           =   2520
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   465
      Left            =   3180
      TabIndex        =   5
      Text            =   "1"
      Top             =   6090
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   6510
      Top             =   6900
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   465
      Index           =   1
      Left            =   3780
      MouseIcon       =   "frmComerciar.frx":0000
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6090
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   540
      MouseIcon       =   "frmComerciar.frx":0152
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6090
      Width           =   2580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MaxHit"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3870
      TabIndex        =   4
      Top             =   870
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MinHit"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3870
      TabIndex        =   3
      Top             =   1110
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2700
      TabIndex        =   2
      Top             =   1140
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   750
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3870
      TabIndex        =   0
      Top             =   630
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   6570
      TabIndex        =   8
      Top             =   6840
      Width           =   315
   End
End
Attribute VB_Name = "frmComerciar"
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

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean

Private Sub cantidad_Change()
    If Val(cantidad.Text) < 1 Then
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
        
        If NpcInv.SelectedItem = 0 Then
            If NpcInv.ItemName(NpcInv.SelectedItem) = "" Then
                Exit Sub
            End If
        End If
        
    
        LasActionBuy = True
        If UserGLD >= NpcInv.Valor(NpcInv.SelectedItem) * Val(cantidad) Then
            Call WriteCommerceBuy(NpcInv.SelectedItem, Int(cantidad.Text))
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
   
   Case 1
   
        If Inventario.SelectedItem = 0 Then Exit Sub
        
        If Inventario.SelectedItem = 0 Then
            If Inventario.ItemName(Inventario.SelectedItem) = "" Then
                Exit Sub
            End If
        End If
   
   
        LasActionBuy = False
        If Not Inventario.Equipped(Inventario.SelectedItem) Then
            Call WriteCommerceSell(Inventario.SelectedItem, Int(cantidad.Text))
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
End Select

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


'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\resources\Graphics\BotónComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\resources\Graphics\Botónvender.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub Image2_Click()
    Call WriteCommerceEnd
End Sub

Private Sub NpcPic_Click()

Dim SR As RECT, DR As RECT

SR.left = 0
SR.top = 0
SR.Right = 32
SR.bottom = 32

DR.left = 0
DR.top = 0
DR.Right = 32
DR.bottom = 32
        
        If NpcInv.SelectedItem = 0 Then Exit Sub
        
        Label1(0).Caption = NpcInv.ItemName(NpcInv.SelectedItem)
        Label1(1).Caption = NpcInv.Valor(NpcInv.SelectedItem)
        Label1(2).Caption = NpcInv.Amount(NpcInv.SelectedItem)
        
        Select Case NpcInv.OBJType(NpcInv.SelectedItem)
            Case eObjType.otWeapon
                Label1(3).Caption = "Max Golpe:" & NpcInv.MaxHit(NpcInv.SelectedItem)
                Label1(4).Caption = "Min Golpe:" & NpcInv.MinHit(NpcInv.SelectedItem)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case eObjType.otArmadura
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & NpcInv.Def(NpcInv.SelectedItem)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        
'        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, NpcInv.GrhIndex(NpcInv.SelectedItem), SR, DR)
        
        
'If Label1(2).Caption = 0 Then ' 27/08/2006 - GS > No mostrar imagen ni nada, cuando no ahi nada que mostrar.
'    Label1(3).Visible = False
'    Label1(4).Visible = False
'    Picture1.Visible = False
'Else
'    Picture1.Visible = True
'    Picture1.Refresh
'End If
        
        
End Sub

Private Sub UsuInv_Click()

Dim SR As RECT, DR As RECT

SR.left = 0
SR.top = 0
SR.Right = 32
SR.bottom = 32

DR.left = 0
DR.top = 0
DR.Right = 32
DR.bottom = 32
        
        If Inventario.SelectedItem = 0 Then Exit Sub

        Label1(0).Caption = Inventario.ItemName(Inventario.SelectedItem)
        Label1(1).Caption = "Precio: " & Inventario.Valor(Inventario.SelectedItem)
        Label1(2).Caption = "Cantidad: " & Inventario.Amount(Inventario.SelectedItem)
        
        Select Case Inventario.OBJType(Inventario.SelectedItem)
            Case eObjType.otWeapon
                Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(Inventario.SelectedItem)
                Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(Inventario.SelectedItem)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case eObjType.otArmadura
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.Def(Inventario.SelectedItem)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        
'        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, Inventario.GrhIndex(Inventario.SelectedItem), SR, DR)
        
'If Label1(2).Caption = 0 Then ' 27/08/2006 - GS > No mostrar imagen ni nada, cuando no ahi nada que mostrar.
'    Label1(3).Visible = False
'    Label1(4).Visible = False
'    Picture1.Visible = False
'Else
'    Picture1.Visible = True
'    Picture1.Refresh
'End If
        
End Sub
