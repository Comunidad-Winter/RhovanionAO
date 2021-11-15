Attribute VB_Name = "modGeneral"
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

Public iplst As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long

Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\Resources\graphics" & "\"
End Function

Public Function DirSound() As String
    DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function

Sub CargarAnimArmas()
On Error Resume Next
    
    Dim loopc As Long
    Dim arch As String
    
    arch = resource_path & PATH_INIT & "\armas.dat"
    
    'NumWeaponAnims = Val(General_Var_Get(arch, "INIT", "NumArmas"))
    
    'ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    'For loopc = 1 To NumWeaponAnims
    '    InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(General_Var_Get(arch, "ARMA" & loopc, "Dir1")), 0
    '    InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(General_Var_Get(arch, "ARMA" & loopc, "Dir2")), 0
    '    InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(General_Var_Get(arch, "ARMA" & loopc, "Dir3")), 0
    '    InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(General_Var_Get(arch, "ARMA" & loopc, "Dir4")), 0
    'Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorH:

    Versiones(1) = Val(General_Var_Get(resource_path & PATH_INIT & "\versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(General_Var_Get(resource_path & PATH_INIT & "\versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(General_Var_Get(resource_path & PATH_INIT & "\versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(General_Var_Get(resource_path & PATH_INIT & "\versiones.ini", "Init", "Val"))
    Versiones(5) = Val(General_Var_Get(resource_path & PATH_INIT & "\versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(General_Var_Get(resource_path & PATH_INIT & "\versiones.ini", "E", "Val"))
    Versiones(7) = Val(General_Var_Get(resource_path & PATH_INIT & "\versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = resource_path & PATH_INIT & "\colores.dat"
    
    If Not General_File_Exists(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).R = CByte(General_Var_Get(archivoC, CStr(i), "R"))
        ColoresPJ(i).G = CByte(General_Var_Get(archivoC, CStr(i), "G"))
        ColoresPJ(i).B = CByte(General_Var_Get(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(51).R = CByte(General_Var_Get(archivoC, "N", "R"))
    ColoresPJ(51).G = CByte(General_Var_Get(archivoC, "N", "G"))
    ColoresPJ(51).B = CByte(General_Var_Get(archivoC, "N", "B"))
    ColoresPJ(50).R = CByte(General_Var_Get(archivoC, "CR", "R"))
    ColoresPJ(50).G = CByte(General_Var_Get(archivoC, "CR", "G"))
    ColoresPJ(50).B = CByte(General_Var_Get(archivoC, "CR", "B"))
    ColoresPJ(49).R = CByte(General_Var_Get(archivoC, "CI", "R"))
    ColoresPJ(49).G = CByte(General_Var_Get(archivoC, "CI", "G"))
    ColoresPJ(49).B = CByte(General_Var_Get(archivoC, "CI", "B"))
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = General_Random_Number(1, 7368)
    CualMITemp = General_Random_Number(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = resource_path & PATH_INIT & "\escudos.dat"
    
    'NumEscudosAnims = Val(General_Var_Get(arch, "INIT", "NumEscudos"))
    
    'ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    'For loopc = 1 To NumEscudosAnims
    '    InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(General_Var_Get(arch, "ESC" & loopc, "Dir1")), 0
    '    InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(General_Var_Get(arch, "ESC" & loopc, "Dir2")), 0
    '    InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(General_Var_Get(arch, "ESC" & loopc, "Dir3")), 0
    '    InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(General_Var_Get(arch, "ESC" & loopc, "Dir4")), 0
    'Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refreshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
         RichTextBox.Refresh 'Le sacamos el refresh, dice nico que con trasparencia no va.
    End With
End Sub


Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(Mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(Mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(Mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

#If SeguridadAlkon Then
    Call UnprotectForm
#End If

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    frmMain.tCheat.Enabled = True
    TiempoActual = GetTickCount
    frmMain.FirstTime = False
    frmMain.tAntiEngine.Enabled = True
    
    Call SaveGameini
    
    'Unload the connect form
    Unload frmPasswd
    Unload frmCrearPersonaje
    Unload frmConnect
    
    'Cuando conecta agregamos lo del Msn, blizzard
    Call SetMusicInfo("Jugando SAAO " & ":" & UserName & " -" & " www.ao.dveloping.com.ar", "", "", "Games", "{1}{0}")

    frmMain.lblItemName.Caption = ""
    
    frmMain.label8.Caption = UserName
    'Load main form
    frmMain.Visible = True
#If SeguridadAlkon Then
    'Unprotect character creation and protect the main form
    Call UnprotectForm
    Call ProtectForm(frmMain)
#End If

End Sub

Sub CargarTip()
    Dim N As Integer
    N = General_Random_Number(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
    
    If Cartel Then Cartel = False
        
    If Engine.Map_Legal_Char_Pos_By_Heading(Engine.User_Char_Index_Get, Direccion) Then
        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            Call WriteWalk(Direccion)
            Call Engine.Char_Move(Engine.User_Char_Index_Get, Direccion)
            Call Engine.Engine_View_Move(Direccion)
        Else
            If UserDescansar And Not UserAvisado Then
                UserAvisado = True
                Call WriteRest
            End If
            If UserMeditar And Not UserAvisado Then
                UserAvisado = True
                Call WriteMeditate
            End If
        End If
    Else
        If Engine.Char_Heading_Get(Engine.User_Char_Index_Get) <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then frmMain.DesactivarMacroTrabajo
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    MoveTo General_Random_Number(NORTH, WEST)
End Sub



Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open resource_path & PATH_INIT & "\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub
Public Function CurServerIp() As String
    CurServerIp = "localhost" '200.58.115.162
End Function

Public Function CurServerPort() As Integer
    CurServerPort = 7666
End Function

Sub Main()
    
    '*******************************************************************************
    '*******************************************************************************
    'Set the resource path.
    resource_path = App.Path & PATH_RESOURCES
    
    'CARGAR CONFIGURACION 'HAY UN MONTON DE COSAS DE AK QUE VUELAN!! INCLUIDA ESTA LINEA!!!
    Call WriteClientVer
    
    'Load config file
    If General_File_Exists(resource_path & PATH_INIT & "\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
        tipf = Config_Inicio.tip
    End If
    'Load ao.dat config file
    If General_File_Exists(resource_path & PATH_INIT & "\ao.dat", vbArchive) Then
        Call LoadClientSetup
    End If
    
    'Cargamos la lista de cheats
    Call LoadCheats

    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.Path & "\")
    
    ChDrive App.Path
    ChDir App.Path
    
    'Obtener HashMD5
    MD5HushYo = MD5File(App.Path & "\SAAO.exe")
    'MD5HushYo = txtOffset(hexMd52Asc(MD5HushYo), 55)
    '*******************************************************************************
    '*******************************************************************************
    
    frmCargando.Show
    frmCargando.Refresh
    
    frmConnect.Version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    
    frmCargando.Label1.Caption = "Iniciando constantes..."
    
    Call InicializarNombres
    ' Initialize FONTTYPES
    Call modProtocol.InitFonts

    
    frmCargando.Label1.Caption = "Finalizado."
    
    frmCargando.Label1.Caption = "Cargando Sonidos..."
    frmCargando.Label1.Caption = "Finalizado"


    Dim loopc As Integer
    
    
    '*******************************************************************************
    '*******************************************************************************
    'INITIALIZATION
    Dim typDevM As typDevMODE
    Dim lRes As Long
    Dim OldResWidth As Integer
    Dim OldResHeight As Integer
    lRes = EnumDisplaySettings(0, 0, typDevM)
    With typDevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        OldResWidth = .dmPelsWidth
        OldResHeight = .dmPelsWidth
        .dmPelsWidth = 800
        .dmPelsHeight = 600
    End With
    'lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    
    'Engine initialization
    If Not DXEngine_Initialize(frmMain.hwnd, frmMain.MainViewPic.hwnd, False) Then
        MsgBox "Se ha producido un error al inicializar el motor grafico, reinstale la aplicacion, si el problema persiste consultar a algun administrador."
    End If
    
    Call modGrh.Animations_Initialize(0.03, 32)
    
    'TileEngine initialization
    If Not Engine.TileEngine_Initialize(False, 7, 146, 17, 13) Then
        MsgBox "Se ha producido un error al inicializar el motor grafico, reinstale la aplicacion, si el problema persiste consultar a algun administrador."
    End If
    
    'Initialize Sound Engine
    Sound_Init
    
    'Initialize Meteorologic
    Call Meteo.Initialize
    
    'Dialogs font
    Call Dialogos.Dialogos_SetFontInfo(1, 8)
    
    'This must be set before rendering...
    Engine.Engine_Base_Speed_Set 0.017 'Speed that the engine should appear to run at
    
    frmCargando.Label1.Caption = "Cargando animaciones extra..."
    
    'Call CargarAnimsExtra
    'Call CargarTips

UserMap = 1

    'Call CargarArrayLluvia
    'Call CargarVersiones
    Call CargarColores
    
#If SeguridadAlkon Then
    CualMI = 0
    Call InitMI
#End If


    frmCargando.Label1.Caption = "Bienvenido a Shadows of Angmar AO."
    
    'Inicializamos el sonido
    frmCargando.Label1.Caption = "Iniciando DirectSound..."
    'Call Audio.Initialize(DirectX, frmMain.hWnd, App.Path & "\" & Config_Inicio.DirSonidos & "\", App.Path & "\" & Config_Inicio.DirMusica & "\")
    
    'Enable / Disable audio
    'Audio.MusicActivated = Not ClientSetup.bNoMusic
    'Audio.SoundActivated = Not ClientSetup.bNoSound
    
    frmCargando.Label1.Caption = "Finalizado..."
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv, MAX_INVENTORY_SLOTS)
    Call NpcInv.Initialize(frmComerciar.NpcPic, MAX_NPC_INVENTORY_SLOTS)
    
    Call modSound.Music_Play(MIdi_Inicio & ".mid")
    Unload frmCargando
    
    frmPres.Picture = LoadPicture(DirGraficos & "alkonao2.jpg")
    frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    
#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

    frmConnect.Visible = True
    'frmConnect.MouseIcon = LoadPicture(App.Path & "\GRAFICOS\Icons\Espada.ico")
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.AttackSpell, INT_ATTACKSPELL)
    Call MainTimer.SetInterval(TimersIndex.Minute, 60000)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.AttackSpell)
    Call MainTimer.Start(TimersIndex.Minute)
    ' Load the form for screenshots
    Call Load(frmScreenshots)
    
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Game_Render
            Meteo.Meteo_Check

            If Not pausa And frmMain.Visible And Not frmForo.Visible And Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmBancoObj.Visible Then
                Game_CheckKeys
            End If
        Else
            Render_Inventory = True
        End If
        
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        DoEvents
    Loop

    'Volvemos a la antigua resolucion
    With typDevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        .dmPelsWidth = OldResWidth
        .dmPelsHeight = OldResHeight
    End With
    
    'lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    'Unload the form for screenshots
    Unload frmScreenshots
    
    'EngineRun = False
    frmCargando.Show
    frmCargando.Label1.Caption = "Liberando Recursos..."

    Call DXEngine_Deinitialize
    
    Set Engine = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    'Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    
    Call EscribirGameIni(Config_Inicio)
    End
End Sub



'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(Mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function
Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim i As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i
End Sub
Public Sub SaveClientSetup()

Dim file As String
Dim nfile As Integer

nfile = FreeFile
file = resource_path & PATH_INIT & "\Setup.ini"



'Cargamos los nuevos datos en el Type ClientSetup
ClientSetup.bDinamic = True
ClientSetup.bNoRes = False
ClientSetup.bUseVideo = False
ClientSetup.byMemory = 40

Open file For Binary As nfile
    Put #nfile, , ClientSetup
Close nfile

End Sub
Public Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 24/06/2006
'
'**************************************************************

    
    'Open App.Path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
    '    Get fHandle, , ClientSetup
    'Close fHandle
    Dim file As String
    Dim fHandle As Integer
    
    fHandle = FreeFile
    file = resource_path & PATH_INIT & "\Setup.ini"
    'Sacamos la info del archivo Setup.ini
    Open file For Binary As fHandle
        Get #fHandle, , ClientSetup
    Close fHandle

    ClientSetup.bDinamic = True
    ClientSetup.bNoRes = False
    ClientSetup.bUseVideo = False
    ClientSetup.byMemory = 40
    
    If ClientSetup.FreeFPS Then
        ClientSetup.FrameInterval = 10
    Else
        ClientSetup.FrameInterval = 56
    End If
    
    'Debug.Print ClientSetup.bDinamic
    'Debug.Print ClientSetup.byMemory
    
    NoRes = ClientSetup.bNoRes
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Arshad"
    'Ciudades(eCiudad.cArghal) = "Arghal"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.Orco) = "Orco"
    
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Fisher) = "Pescador"
    ListaClases(eClass.Blacksmith) = "Herrero"
    ListaClases(eClass.Lumberjack) = "Leñador"
    ListaClases(eClass.Miner) = "Minero"
    ListaClases(eClass.Carpenter) = "Carpintero"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Suerte) = "Suerte"
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
    SkillsNames(eSkill.armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
    SkillsNames(eSkill.Wresterling) = "Wrestling"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.BorrarDialogos
End Sub

Public Function General_File_Exists(ByVal file_path As String, ByVal file_type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************
    If Dir(file_path, file_type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal Value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, Value, file
End Sub

Public Function general_field_read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(Mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                general_field_read = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        general_field_read = Mid$(Text, LastPos + 1)
    End If
End Function

Public Function General_Field_Count(ByVal Text As String, ByVal delimiter As Byte) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Count the number of fields in a delimited string
'*****************************************************************
    'If string is empty there aren't any fields
    If Len(Text) = 0 Then
        Exit Function
    End If

    Dim i As Long
    Dim FieldNum As Long
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(Mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
        End If
    Next i
    General_Field_Count = FieldNum + 1
End Function

Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
'*****************************************************************
'Author: Aaron Perkins
'Find a Random number between a range
'*****************************************************************
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Sub General_Sleep(ByVal length As Double)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sleep for a given number a seconds
'*****************************************************************
    Dim curFreq As Currency
    Dim curStart As Currency
    Dim curEnd As Currency
    Dim dblResult As Double
    
    QueryPerformanceFrequency curFreq 'Get the timer frequency
    QueryPerformanceCounter curStart 'Get the start time
    
    Do Until dblResult >= length
        QueryPerformanceCounter curEnd 'Get the end time
        dblResult = (curEnd - curStart) / curFreq 'Calculate the duration (in seconds)
    Loop
End Sub

Public Function General_Get_Elapsed_Time() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If

    'Get current time
    QueryPerformanceCounter start_time
    
    'Calculate elapsed time
    General_Get_Elapsed_Time = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    QueryPerformanceCounter end_time
End Function

Public Function General_Get_Free_Ram_Bytes() As Long
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys
End Function

Public Sub ActualizarCoordenadas()
    Dim x As Long
    Dim y As Long
    
    Call Engine.Char_Map_Pos_Get(Engine.User_Char_Index_Get, x, y)
    frmMain.Coord.Caption = "Map:" & UserMap & " X:" & x & " Y:" & y
End Sub

Public Sub General_Screen_Left_Click(ByVal tX As Long, ByVal tY As Long, Optional ByVal Shift As Byte = 0)
    If Comerciando Then Exit Sub
    Select Case Shift
        Case 0
            If UsaMacro Then
                CnTd = CnTd + 1
                If CnTd = 3 Then
                    Call WriteUseSpellMacro
                    CnTd = 0
                End If
                UsaMacro = False
            End If
            If UsingSkill = 0 Then
                Call WriteLeftClick(tX, tY)
            Else
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.macrotrabajo.Enabled Then frmMain.DesactivarMacroTrabajo
                'No mas intervalo Hechi-Golpe?
                If Not MainTimer.Check(TimersIndex.AttackSpell) Then Exit Sub              'Check if attack interval has finished.
                If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub                                    'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then Exit Sub 'Check if spells interval has finished.
                'Splitted because VB isn't lazy!
                If UsingSkill = Proyectiles Then
                    If Not MainTimer.Check(TimersIndex.Arrows) Then
                        Exit Sub
                    End If
                End If
                'Splitted because VB isn't lazy!
                If UsingSkill = Magia Then
                    If Not MainTimer.Check(TimersIndex.CastSpell) Then
                        Exit Sub
                    End If
                End If
                'Splitted because VB isn't lazy!
                If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                    If Not MainTimer.Check(TimersIndex.Work) Then
                        Exit Sub
                    End If
                End If
                
                frmMain.MousePointer = vbDefault
                Call WriteWorkLeftClick(tX, tY, UsingSkill)
                UsingSkill = 0
            End If
        Case 1
            Call WriteWarpChar("YO", UserMap, tX, tY)
    End Select
End Sub

Public Sub General_Screen_Double_Click(ByVal tX As Long, ByVal tY As Long, Optional ByVal Shift As Byte = 0)
    If Not frmForo.Visible Then
        Call WriteDoubleClick(tX, tY)
    End If
End Sub

Public Sub SpeedCalculate(ByVal timer_elapsed_time As Single)
    Engine.TileEngineSpeedCalculate timer_elapsed_time
    ParticleSpeedCalculate timer_elapsed_time
    AnimSpeedCalculate timer_elapsed_time
End Sub
