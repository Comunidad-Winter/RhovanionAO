Attribute VB_Name = "modGame"
'EXTERNAL FUNCTIONS
'KeyInput
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
'time between frames
Dim timer_elapsed_time As Single


'***********************
'CONSTATNS
'***********************
'Objetos


'***********************
'Type
'***********************


'***********************
'Enums
'***********************



Public Function Input_Key_Get(ByVal key_code As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    If GetKeyState(key_code) < 0 Then
        Input_Key_Get = True
    End If
End Function

Public Sub Game_Render()
    
    '*******************
    'RenderScreen
    '*******************
    DXEngine_BeginRender
        
    '*******************************
    'Draw Map
    Engine.Map_Render
    '*******************************
    
    '*******************************
    'RenderSignal
    DibujarCartel
    '*******************************
    
    '*******************************
    'Render Dialogs
    Dialogos.MostrarTexto
    '*******************************
    
    '*******************************
    'Draw engine stats
    DXEngine_StatsRender
    '*******************************
    
    DXEngine_EndRender
    
    
    '*******************
    'RenderInventory
    '*******************
    If Render_Inventory Then
        Render_Inventory = False
        If frmComerciar.Visible Then
            DXEngine_BeginSecondaryRender
            NpcInv.DrawInventory
            DXEngine_EndSecondaryRender frmComerciar.NpcPic.hwnd, 168, 269
            DXEngine_BeginSecondaryRender
            Inventario.DrawInventory
            DXEngine_EndSecondaryRender frmComerciar.UsuInv.hwnd, 168, 269
        ElseIf frmBancoObj.Visible Then
            DXEngine_BeginSecondaryRender
            NpcInv.DrawInventory
            DXEngine_EndSecondaryRender frmBancoObj.InvNpc.hwnd, 168, 269
            DXEngine_BeginSecondaryRender
            Inventario.DrawInventory
            DXEngine_EndSecondaryRender frmBancoObj.invUsu.hwnd, 168, 269
        Else
            DXEngine_BeginSecondaryRender
            Inventario.DrawInventory
            DXEngine_EndSecondaryRender frmMain.picInv.hwnd, 162, 166
        End If
    End If
    
    timer_elapsed_time = General_Get_Elapsed_Time()
    SpeedCalculate (timer_elapsed_time)
End Sub
Public Sub Game_CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'No input allowed while Argentum is not the active window
    If Not modApi.IsAppActive() Then Exit Sub
    If Not frmMain.Visible Then Exit Sub
    
    'Dont allow pressing this keys if we are moving
    If Not Engine.Player_Moving Then
        If Not UserEstupido Then
            If Input_Key_Get(vbKeyUp) Then
                Call MoveTo(E_Heading.NORTH)
            ElseIf Input_Key_Get(vbKeyRight) Then
                Call MoveTo(E_Heading.EAST)
            ElseIf Input_Key_Get(vbKeyDown) Then
                Call MoveTo(E_Heading.SOUTH)
            ElseIf Input_Key_Get(vbKeyLeft) Then
                Call MoveTo(E_Heading.WEST)
            End If
        Else
            If Input_Key_Get(vbKeyRight) Or Input_Key_Get(vbKeyLeft) Or Input_Key_Get(vbKeyUp) Or Input_Key_Get(vbKeyDown) Then
                Call RandomMove 'Si presiona cualquier tecla y es estupido se mueve para cualquier lado.
            End If
        End If
        Call ActualizarCoordenadas
    End If
End Sub


