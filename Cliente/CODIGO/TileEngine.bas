Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez



Option Explicit

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'    C       O       N       S      T
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'    T       I      P      O      S
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'Encabezado bmp
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    x As Integer
    y As Integer
End Type

Public Type OffSet
    x As Single
    y As Single
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tama�o y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Single
    'Active As Boolean
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    SpeedCounter As Single
    Started As Byte
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As OffSet
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type


'Lista de cuerpos
Public Type FxData
    fX As Integer
    OffSetX As Long
    OffSetY As Long
End Type

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As Byte ' As E_Heading ?
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    fX As Integer
    fXGrh As Grh
    FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    
    Moving As Byte
    MoveOffset As OffSet
    ServerIndex As Integer
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type


Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tama�o del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tama�o muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tama�o de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer


'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim engineBaseSpeed As Single
Dim UserSpeed As Single
'?�?�?�?�?�?�?�?�?�?�Totales?�?�?�?�?�?�?�?�?�?�?

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

#If ConAlfaB Then
    Public MotorG As New Cls_Motor
#End If

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'�?�?�?�?�?�?�?�?�?�Graficos�?�?�?�?�?�?�?�?�?�?�?

Public LastTime As Long 'Para controlar la velocidad


'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer




'�?�?�?�?�?�?�?�?�?�Graficos�?�?�?�?�?�?�?�?�?�?�?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh 'Animaciones publicas
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�Mapa?�?�?�?�?�?�?�?�?�?�?�?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�Usuarios?�?�?�?�?�?�?�?�?�?�?�?�?
'
'epa ;)
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�?�API?�?�?�?�?�?�?�?�?�?�?�?�?�?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'GetElapsedTime
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long


'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'est� raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long
Public DeNoche      As Byte 'Nochee!
Public bMapa        As Boolean

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

#If ConAlfaB Then

Public AlphaTechos As Byte

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long

Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, _
     ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, _
     ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer

Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
     ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
     ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
     ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
     ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

#End If

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long




Sub CargarCabezas()
On Error Resume Next
Dim N As Integer, i As Integer, Numheads As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Numheads

'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #N, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #N

End Sub

Sub CargarCascos()
On Error Resume Next
Dim N As Integer, i As Integer, NumCascos As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #N, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #N

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

N = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #N, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
Next i

Close #N

End Sub
Sub CargarFxs()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

N = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumFxs

'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #N, , MisFxs(i)
    FxData(i).fX = MisFxs(i).Animacion
    FxData(i).OffSetX = MisFxs(i).OffSetX
    FxData(i).OffSetY = MisFxs(i).OffSetY
Next i

Close #N

End Sub

Sub CargarTips()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumTips As Integer

N = FreeFile
Open App.Path & "\init\Tips.ayu" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumTips

'Resize array
ReDim Tips(1 To NumTips) As String * 255

For i = 1 To NumTips
    Get #N, , Tips(i)
Next i

Close #N

End Sub

Sub CargarArrayLluvia()
On Error Resume Next
Dim N As Integer, i As Integer
Dim Nu As Integer

N = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Nu

'Resize array
ReDim bLluvia(1 To Nu) As Byte

For i = 1 To Nu
    Get #N, , bLluvia(i)
Next i

Close #N

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal cx As Single, ByVal cy As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

cx = cx - StartPixelLeft
cy = cy - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
cx = (cx \ TilePixelWidth)
cy = (cy \ TilePixelHeight)

If cx > HWindowX Then
    cx = (cx - HWindowX)

Else
    If cx < HWindowX Then
        cx = (0 - (HWindowX - cx))
    Else
        cx = 0
    End If
End If

If cy > HWindowY Then
    cy = (0 - (HWindowY - cy))
Else
    If cy < HWindowY Then
        cy = (cy - HWindowY)
    Else
        cy = 0
    End If
End If

tX = UserPos.x + cx
tY = UserPos.y + cy

End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

'Apuntamos al ultimo Char
If CharIndex > LastChar Then LastChar = CharIndex

'If the char wasn't allready active (we are rewritting it) don't increase char count
If charlist(CharIndex).Active = 0 Then _
    NumChars = NumChars + 1

If Arma = 0 Then Arma = 2
If Escudo = 0 Then Escudo = 2
If Casco = 0 Then Casco = 2

charlist(CharIndex).iHead = Head
charlist(CharIndex).iBody = Body
charlist(CharIndex).Head = HeadData(Head)
charlist(CharIndex).Body = BodyData(Body)
charlist(CharIndex).Arma = WeaponAnimData(Arma)
'[ANIM ATAK]
charlist(CharIndex).Arma.WeaponAttack = 0

charlist(CharIndex).Escudo = ShieldAnimData(Escudo)
charlist(CharIndex).Casco = CascoAnimData(Casco)

charlist(CharIndex).Heading = Heading

'Reset moving stats
charlist(CharIndex).Moving = 0
charlist(CharIndex).MoveOffset.x = 0
charlist(CharIndex).MoveOffset.y = 0

'Update position
charlist(CharIndex).Pos.x = x
charlist(CharIndex).Pos.y = y

'Make active
charlist(CharIndex).Active = 1

'Plot on map
MapData(x, y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    charlist(CharIndex).Active = 0
    charlist(CharIndex).Criminal = 0
    charlist(CharIndex).fX = 0
    charlist(CharIndex).FxLoopTimes = 0
    charlist(CharIndex).invisible = False

#If SeguridadAlkon Then
    Call MI(CualMI).ResetInvisible(CharIndex)
#End If

    charlist(CharIndex).Moving = 0
    charlist(CharIndex).muerto = False
    charlist(CharIndex).Nombre = ""
    charlist(CharIndex).pie = False
    charlist(CharIndex).Pos.x = 0
    charlist(CharIndex).Pos.y = 0
    charlist(CharIndex).UsandoArma = False
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

charlist(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until charlist(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).CharIndex = 0

Call ResetCharInfo(CharIndex)

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1
'[CODE 000]:MatuX
'
'  La linea generaba un error en la IDE, (no ocurr�a debido al
' on error)
'
'   Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
'[END]'

End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim x As Integer
Dim y As Integer
Dim nX As Integer
Dim nY As Integer

x = charlist(CharIndex).Pos.x
y = charlist(CharIndex).Pos.y

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        addY = -1

    Case E_Heading.EAST
        addX = 1

    Case E_Heading.SOUTH
        addY = 1
    
    Case E_Heading.WEST
        addX = -1
        
End Select

nX = x + addX
nY = y + addY

MapData(nX, nY).CharIndex = CharIndex
charlist(CharIndex).Pos.x = nX
charlist(CharIndex).Pos.y = nY
MapData(x, y).CharIndex = 0

charlist(CharIndex).MoveOffset.x = -1 * (TilePixelWidth * addX)
charlist(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)

charlist(CharIndex).scrollDirectionX = addX
charlist(CharIndex).scrollDirectionY = addY

charlist(CharIndex).Moving = 1
charlist(CharIndex).Heading = nHeading

If UserEstado <> 1 Then Call DoPasosFx(CharIndex)

'areas viejos
If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(CharIndex)
End If

End Sub

Public Sub DoFogataFx()
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata()
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
    End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean
    Dim x As Long, y As Long
    
    For y = UserPos.y - MinYBorder + 1 To UserPos.y + MinYBorder - 1
        For x = UserPos.x - MinXBorder + 1 To UserPos.x + MinXBorder - 1
            If MapData(x, y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        Next x
    Next y
    
    EstaPCarea = False
End Function


Sub DoPasosFx(ByVal CharIndex As Integer)
Static pie As Boolean

If Not UserNavegando Then
    If Not charlist(CharIndex).muerto And EstaPCarea(CharIndex) And (charlist(CharIndex).priv = 0 Or charlist(CharIndex).priv > 5) Then
        charlist(CharIndex).pie = Not charlist(CharIndex).pie
        If charlist(CharIndex).pie Then
            Call Audio.PlayWave(SND_PASOS1)
        Else
            Call Audio.PlayWave(SND_PASOS2)
        End If
    End If
Else
    Call Audio.PlayWave(SND_NAVEGANDO)
End If

End Sub


Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

On Error Resume Next

Dim x As Integer
Dim y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As E_Heading



x = charlist(CharIndex).Pos.x
y = charlist(CharIndex).Pos.y

MapData(x, y).CharIndex = 0

addX = nX - x
addY = nY - y

If Sgn(addX) = 1 Then
    nHeading = E_Heading.EAST
End If

If Sgn(addX) = -1 Then
    nHeading = E_Heading.WEST
End If

If Sgn(addY) = -1 Then
    nHeading = E_Heading.NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = E_Heading.SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex


charlist(CharIndex).Pos.x = nX
charlist(CharIndex).Pos.y = nY

charlist(CharIndex).MoveOffset.x = -1 * (TilePixelWidth * addX)
charlist(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)

charlist(CharIndex).scrollDirectionX = Sgn(addX)
charlist(CharIndex).scrollDirectionY = Sgn(addY)

charlist(CharIndex).Moving = 1
charlist(CharIndex).Heading = nHeading

'parche para que no medite cuando camina
Dim fxCh As Integer
fxCh = charlist(CharIndex).fX
If fxCh = FxMeditar.CHICO Or fxCh = FxMeditar.GRANDE Or fxCh = FxMeditar.MEDIANO Or fxCh = FxMeditar.XGRANDE Then
    charlist(CharIndex).fX = 0
    charlist(CharIndex).FxLoopTimes = 0
End If

If Not EstaPCarea(CharIndex) Then Call Dialogos.QuitarDialogo(CharIndex)

If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(CharIndex)
End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim x As Integer
Dim y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        y = -1

    Case E_Heading.EAST
        x = 1

    Case E_Heading.SOUTH
        y = 1
    
    Case E_Heading.WEST
        x = -1
        
End Select

'Fill temp pos
tX = UserPos.x + x
tY = UserPos.y + y

'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.x = x
    UserPos.x = tX
    AddtoUserPos.y = y
    UserPos.y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.x - 8 To UserPos.x + 8
    For k = UserPos.y - 6 To UserPos.y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopc As Integer
Dim Dale As Boolean

loopc = 1
Do While charlist(loopc).Active And Dale
    loopc = loopc + 1
    Dale = (loopc <= UBound(charlist))
Loop

NextOpenChar = loopc

End Function


Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer




'Resize arrays
ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

'Open files
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
    
    'GrhData(Grh).Active = True
    
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh & " " & Err.Description

End Sub

Function LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Limites del mapa
If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    'Tile Bloqueado?
    If MapData(x, y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    '�Hay un personaje?
    If MapData(x, y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(x, y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(x, y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function




Function InMapLegalBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte)

'Dim CurrentGrh As Grh
'Dim destRect As RECT
'Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2

'If Animate Then
'    If Grh.Started = 1 Then
'        If Grh.SpeedCounter > 0 Then
'            Grh.SpeedCounter = Grh.SpeedCounter - 1
'            If Grh.SpeedCounter = 0 Then
'                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'                Grh.FrameCounter = Grh.FrameCounter + 1
'                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
'                    Grh.FrameCounter = 1
'                End If
'            End If
'        End If
'    End If
'End If
'Figure out what frame to draw (always 1 if not animated)
'CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
'Center Grh over X,Y pos
'If center Then
'    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
'        x = x - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
'    End If
'    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
'        y = y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
'    End If
'End If
'With SourceRect
'        .Left = GrhData(CurrentGrh.GrhIndex).sX
'        .Top = GrhData(CurrentGrh.GrhIndex).sY
'        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
'        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
'End With
'Surface.BltFast x, y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT

Dim CurrentGrhIndex As Integer
Dim SourceRect As RECT
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                'If Grh.Loops <> INFINITE_LOOPS Then
                '    If Grh.Loops > 0 Then
                '        Grh.Loops = Grh.Loops - 1
                '    Else
                '        Grh.Started = 0
                '    End If
                'End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
        
        If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With SourceRect
        .Left = GrhData(CurrentGrhIndex).sX
        .Top = GrhData(CurrentGrhIndex).sY
        .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
    End With
    
    'Draw
    Call Surface.BltFast(x, y, SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT)
End Sub

Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, Grh As Integer, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = x
    .Top = y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.Surface(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

'Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[CODE 000]:MatuX
    Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional Capa2 As Boolean = False, Optional WeaponAttack As Byte = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
'Dim iGrhIndex As Integer
'Dim destRect As RECT
'Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
'Dim QuitarAnimacion As Boolean


'If Animate Then
'    If Grh.Started = 1 Then
'        If Grh.SpeedCounter > 0 Then
'            Grh.SpeedCounter = Grh.SpeedCounter - 1
'            If Grh.SpeedCounter = 0 Then
'                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'                Grh.FrameCounter = Grh.FrameCounter + 1
'                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
'                    Grh.FrameCounter = 1
'                    If KillAnim Then
'                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
'
'                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
'                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
'                                charlist(KillAnim).fX = 0
'                                Exit Sub
'                            End If
                            
'                        End If
'                    End If
'               End If
'            End If
'        End If
'    End If
'End If

    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter) '(timerTicksPerFrame * Grh.SpeedCounter) '(timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)'timerTicksPerFrame * (Grh.SpeedCounter / 1000) '(timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter) '(timerTicksPerFrame * Grh.SpeedCounter) '(timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter) 'Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                If WeaponAttack > 0 Then
                    WeaponAttack = 0
                End If
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                'If Grh.Loops <> INFINITE_LOOPS Then
                '    If Grh.Loops > 0 Then
                '        Grh.Loops = Grh.Loops - 1
                '    Else
                '        Grh.Started = 0
                '    End If
                'End If
                If KillAnim Then
                    If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                        If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                        If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                            charlist(KillAnim).fX = 0
                            Exit Sub
                        End If
                            
                    End If
                End If
            End If
        End If
    End If

    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
            
        If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With SourceRect
        .Left = GrhData(CurrentGrhIndex).sX
        .Top = GrhData(CurrentGrhIndex).sY
        .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
    End With
    

'If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
'iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

''Center Grh over X,Y pos
'If center Then
'    If GrhData(iGrhIndex).TileWidth <> 1 Then
'        If Capa2 = True Then ' [GS] 24/10/2006 - Correccion en la capa 2
'            x = x - Int(GrhData(iGrhIndex).TileWidth * 32) + 32 'hard coded for speed
'        Else
'            x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
'        End If
'    End If
'    If GrhData(iGrhIndex).TileHeight <> 1 Then
'        y = y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
'    End If
'End If

'With SourceRect
'    .Left = GrhData(iGrhIndex).sX
'    .Top = GrhData(iGrhIndex).sY
'    .Right = .Left + GrhData(iGrhIndex).pixelWidth
'    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
'End With


Surface.BltFast x, y, SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub

'Saao Minimap.

Public Sub DibujarMiniMapa(ByRef Pic As PictureBox)
    Dim DR As RECT
    DR.Left = 0
    DR.Top = 0
    DR.Bottom = 100
    DR.Right = 100
    SupMiniMap.Blt DR, SupBMiniMap, DR, DDBLT_DONOTWAIT
    DR.Left = UserPos.x - 1
    DR.Top = UserPos.y - 1
    DR.Bottom = UserPos.y + 1
    DR.Right = UserPos.x + 1
    SupMiniMap.BltColorFill DR, vbRed
    DR.Left = 0
    DR.Top = 0
    DR.Bottom = 100
    DR.Right = 100
    SupMiniMap.BltToDC Pic.hdc, DR, DR
End Sub

Public Sub GenerarMiniMapa()
    Dim x As Integer
    Dim y As Integer
    Dim i As Integer
    Dim DR As RECT
    Dim SR As RECT
    Dim aux As Integer
    
    'Dim OffSetX As Byte
    'Dim OffSetY As Byte
    
    SR.Left = 0
    SR.Top = 0
    SR.Bottom = 100
    SR.Right = 100
    'SupBMiniMap.BltColorFill SR, vbBlack
    
    For x = MinYBorder To MaxXBorder
        For y = MinYBorder To MaxYBorder
            If MapData(x, y).Graphic(1).GrhIndex > 0 Then
                With MapData(x, y).Graphic(1)
                    i = GrhData(.GrhIndex).Frames(1)
                End With
                
                SR.Left = GrhData(i).sX
                SR.Top = GrhData(i).sY
                SR.Right = GrhData(i).sX + GrhData(i).pixelWidth
                SR.Bottom = GrhData(i).sY + GrhData(i).pixelHeight
                DR.Left = x
                DR.Top = y
                DR.Bottom = y + 1
                DR.Right = x + 1
                SupBMiniMap.Blt DR, SurfaceDB.Surface(GrhData(i).FileNum), SR, DDBLT_DONOTWAIT
                'SupMiniMap.BltFast x, y, SurfaceDB.GetBMP(GrhData(i).FileNum), SR, DDBLTFAST_DESTCOLORKEY
            End If
            
            If MapData(x, y).Graphic(3).GrhIndex > 0 Then
                With MapData(x, y).Graphic(3)
                    i = GrhData(.GrhIndex).Frames(1)
                End With
            
                SR.Left = GrhData(i).sX
                SR.Top = GrhData(i).sY
                SR.Right = GrhData(i).sX + GrhData(i).pixelWidth
                SR.Bottom = GrhData(i).sY + GrhData(i).pixelHeight
                DR.Left = x
                DR.Top = y
                DR.Bottom = y + 1
                DR.Right = x + 1
                SupBMiniMap.Blt DR, SurfaceDB.Surface(GrhData(i).FileNum), SR, DDBLT_DONOTWAIT
                'SupMiniMap.BltFast x, y, SurfaceDB.GetBMP(GrhData(i).FileNum), SR, DDBLTFAST_DESTCOLORKEY
            End If
        Next
    Next
    
End Sub
#If ConAlfaB Then
Public Sub DDrawAlGrhtoSurface(ByRef Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional AlphaBit As Byte = 150)

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean

Dim ddsdSrc As DDSURFACEDESC2
Dim Modo As Long

Surface.GetSurfaceDesc ddsdSrc


If ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
End If

'If Modo <> 1 Then Exit Sub

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).fX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
MotorG.DBAlpha SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, True, x, y, AlphaBit ' AlphaBit
End Sub
#End If


#If ConAlfaB = 1 Then
    Sub DDrawTransGrhtoSurfaceAlpha(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter) '(timerTicksPerFrame * Grh.SpeedCounter) '(timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter) 'Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                'If Grh.Loops <> INFINITE_LOOPS Then
                '    If Grh.Loops > 0 Then
                '        Grh.Loops = Grh.Loops - 1
                '    Else
                '        Grh.Started = 0
                '    End If
                'End If
                If KillAnim Then
                    If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                        If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                        If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                            charlist(KillAnim).fX = 0
                            Exit Sub
                        End If
                            
                    End If
                End If
            End If
        End If
    End If

    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
            
        If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With SourceRect
        .Left = GrhData(CurrentGrhIndex).sX
        .Top = GrhData(CurrentGrhIndex).sY
        .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
    End With



'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set Src = SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum)

Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = x
    .Top = y
    .Right = x + GrhData(CurrentGrhIndex).pixelWidth
    .Bottom = y + GrhData(CurrentGrhIndex).pixelHeight
    
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast x, y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(x + x, y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)

Surface.Unlock rDest
DstLock = False
Src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub
#End If 'ConAlfaB = 1

Sub DrawBackBufferSurface()
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hWnd As Long, hdc As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
    If Grh <= 0 Then Exit Sub
    
    SecundaryClipper.SetHWnd hWnd
    SurfaceDB.Surface(GrhData(Grh).FileNum).BltToDC hdc, SourceRect, destRect
End Sub

Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Single, PixelOffsetY As Single)
'***************************************************
'Autor: Unknown
'Last Modification: 12/24/06
'
'12/24/06: NIGO - check X,Y are in map bounds
'***************************************************
On Error Resume Next


If UserCiego Then Exit Sub

Dim y        As Integer 'Keeps track of where on map we are
Dim x        As Integer 'Keeps track of where on map we are
Dim minY     As Integer 'Start Y pos on current map
Dim maxY     As Integer 'End Y pos on current map
Dim minX     As Integer 'Start X pos on current map
Dim maxX     As Integer 'End X pos on current map
Dim ScreenX  As Integer 'Keeps track of where to place tile on screen
Dim ScreenY  As Integer 'Keeps track of where to place tile on screen
Dim moved    As Byte
Dim Grh      As Grh     'Temp Grh for show tile and blocked
Dim TempChar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer 'Usado en el Layer de Chars
Dim iPPy     As Integer 'Usado en el Layer de Chars
Dim rSourceRect      As RECT    'Usado en el Layer 1
Dim iGrhIndex        As Integer 'Usado en el Layer 1
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim nX As Integer
Dim nY As Integer

'Figure out Ends and Starts of screen
' Hardcodeado para speed!
minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)


'Draw floor layer
ScreenY = 8
For y = (minY + 8) To maxY - 8
    If y > 0 And y < 101 Then 'In map bounds
        ScreenX = 8
        For x = minX + 8 To maxX - 8
            If x > 0 And x < 101 Then 'In map bounds
                'Layer 1 **********************************
                If MapData(x, y).Graphic(1).GrhIndex <> 0 Then
                    Call DDrawGrhtoSurface(BackBufferSurface, MapData(x, y).Graphic(1), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        0, 1)
                End If
                '******************************************
                'Layer 2 **********************************
                If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            MapData(x, y).Graphic(2), _
                            ((32 * ScreenX) - 32) + PixelOffsetX, _
                            ((32 * ScreenY) - 32) + PixelOffsetY, _
                            1, _
                            1, 0, True)
                End If
                '******************************************
            End If
            ScreenX = ScreenX + 1
        Next x
    End If
    ScreenY = ScreenY + 1
Next y


'busco que nombre dibujar
Call ConvertCPtoTP(frmMain.MainViewShp.Left, frmMain.MainViewShp.Top, frmMain.MouseX, frmMain.MouseY, nX, nY)


'Draw Transparent Layers  (Layer 2, 3)
ScreenY = 8
For y = minY + 8 To maxY - 1
    If y > 0 And y < 101 Then 'In map bounds
        ScreenX = 5
        For x = minX + 5 To maxX - 5
            If x > 0 And x < 101 Then 'In map bounds
                iPPx = 32 * ScreenX - 32 + PixelOffsetX
                iPPy = 32 * ScreenY - 32 + PixelOffsetY

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
'                   If Y > UserPos.Y Then
'                       Call DDrawTransGrhtoSurfaceAlpha( _
'                               BackBufferSurface, _
'                               MapData(X, Y).ObjGrh, _
'                               iPPx, iPPy, 1, 1)
'                   Else
                        Call DDrawTransGrhtoSurface( _
                                BackBufferSurface, _
                                MapData(x, y).ObjGrh, _
                                iPPx, iPPy, 1, 1)
'                   End If
                End If
                '***********************************************
                'Char layer ************************************
                If MapData(x, y).CharIndex <> 0 Then
                    TempChar = charlist(MapData(x, y).CharIndex)
                    PixelOffsetXTemp = PixelOffsetX
                    PixelOffsetYTemp = PixelOffsetY
                    moved = 0
            
                    'Dibuja solamente players
                    iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
                    iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
                    
                    Call CharRender(MapData(x, y).CharIndex, iPPx, iPPy)

                End If '<-> If MapData(X, Y).CharIndex <> 0 Then
                '*************************************************
                'Layer 3 *****************************************
                If MapData(x, y).Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            MapData(x, y).Graphic(3), _
                            ((32 * ScreenX) - 32) + PixelOffsetX, _
                            ((32 * ScreenY) - 32) + PixelOffsetY, _
                            1, 1)
                End If
                '************************************************
            End If
            ScreenX = ScreenX + 1
        Next x
    End If
    ScreenY = ScreenY + 1
Next y

If Not bTecho Then
    #If ConAlfaB Then
    If ClientSetup.bFading Then
        If AlphaTechos < 255 Then
            frmMain.ATecho.Enabled = True
        End If
    Else
        AlphaTechos = 255
    End If
    #End If
    'Draw blocked tiles and grid
    ScreenY = 5
    For y = minY + 5 To maxY - 1
        If y > 0 And y < 101 Then 'In map bounds
            ScreenX = 5
            For x = minX + 5 To maxX
                If y > 0 And y < 101 Then 'In map bounds
                    If MapData(x, y).Graphic(4).GrhIndex <> 0 Then
                        'Draw
                        #If ConAlfaB Then
                        If AlphaTechos <> 255 Then
                        Call DDrawAlGrhtoSurface( _
                            BackBufferSurface, _
                            MapData(x, y).Graphic(4), _
                            ((32 * ScreenX) - 32) + PixelOffsetX, _
                            ((32 * ScreenY) - 32) + PixelOffsetY, _
                            1, 1, 0, AlphaTechos)
                        Else
                        #End If
                        Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            MapData(x, y).Graphic(4), _
                            ((32 * ScreenX) - 32) + PixelOffsetX, _
                            ((32 * ScreenY) - 32) + PixelOffsetY, _
                            1, 1, 0)
                        #If ConAlfaB Then
                        End If
                        #End If
                    End If
                End If
                ScreenX = ScreenX + 1
            Next x
        End If
        ScreenY = ScreenY + 1
    Next y
#If ConAlfaB Then
Else

If ClientSetup.bFading Then
    If AlphaTechos >= 80 Then
        frmMain.ATecho.Enabled = True
    End If

ScreenY = 5
    For y = minY + 5 To maxY - 1
        If y > 0 And y < 101 Then 'In map bounds
            ScreenX = 5
            For x = minX + 5 To maxX
                If y > 0 And y < 101 Then 'In map bounds
                    If MapData(x, y).Graphic(4).GrhIndex <> 0 Then
                        'Draw
                        Call DDrawAlGrhtoSurface( _
                            BackBufferSurface, _
                            MapData(x, y).Graphic(4), _
                            ((32 * ScreenX) - 32) + PixelOffsetX, _
                            ((32 * ScreenY) - 32) + PixelOffsetY, _
                            1, 1, 0, AlphaTechos)
                    End If
                End If
                ScreenX = ScreenX + 1
            Next x
        End If
        ScreenY = ScreenY + 1
    Next y
End If
#End If
End If

If bLluvia(UserMap) = 1 Then
    If bRain Then
                'Figure out what frame to draw
                If llTick < DirectX.TickCount - 50 Then
                    iFrameIndex = iFrameIndex + 1
                    If iFrameIndex > 7 Then iFrameIndex = 0
                    llTick = DirectX.TickCount
                End If
    
                For y = 0 To 4
                    For x = 0 To 4
                        Call BackBufferSurface.BltFast(LTLluvia(y), LTLluvia(x), SurfaceDB.Surface(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
                    Next x
                Next y
    End If
End If




Dim PP As RECT

PP.Left = 0
PP.Top = 0
PP.Right = 400
PP.Bottom = 400

'Call BackBufferSurface.BltFast(LTLluvia(0) + TilePixelWidth, LTLluvia(0) + TilePixelHeight, SurfaceDB.surface(10000), PP, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
#If ConAlfaB Then
    If ClientSetup.bNoche Then
        If bLluvia(UserMap) Then
            If DeNoche <> 1 Then
                EfectoNoche BackBufferSurface
            End If
        End If
    End If
#End If

Call DibujarMiniMapa(frmMain.picMiniMap)

If bMapa = True Then
    Call BackBufferSurface.BltFast(340, MainViewRect.Top + 125, SurfaceDB.Surface(26029), PP, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
End If

'[USELESS]:El codigo para llamar a la noche, nublado, etc.
'            If bTecho Then
'                Dim bbarray() As Byte, nnarray() As Byte
'                Dim ddsdBB As DDSURFACEDESC2 'backbuffer
'                Dim ddsdNN As DDSURFACEDESC2 'nnublado
'                Dim r As RECT, r2 As RECT
'                Dim retVal As Long
'                '[LOCK]:BackBufferSurface
'                    BackBufferSurface.GetSurfaceDesc ddsdBB
'                    'BackBufferSurface.Lock r, ddsdBB, DDLOCK_NOSYSLOCK + DDLOCK_WRITEONLY + DDLOCK_WAIT, 0
'                    BackBufferSurface.Lock r, ddsdBB, DDLOCK_WRITEONLY + DDLOCK_WAIT, 0
'                    BackBufferSurface.GetLockedArray bbarray()
''                '[LOCK]:BBMask
''                    SurfaceXU(2).GetSurfaceDesc ddsdNN
''                    'SurfaceXU(2).Lock r2, ddsdNN, DDLOCK_READONLY + DDLOCK_NOSYSLOCK + DDLOCK_WAIT, 0
''                    SurfaceXU(2).Lock r2, ddsdNN, DDLOCK_READONLY + DDLOCK_WAIT, 0
''                    SurfaceXU(2).GetLockedArray nnarray()
'                '[BLIT]'
'                    'retVal = BlitNoche(bbarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth, 0)
'                    'retval = BlitNublar(bbarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth)
'                    'retVal = BlitNublarMMX(bbarray(0, 0), nnarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth, ddsdBB.lPitch, ddsdNN.lHeight, ddsdNN.lWidth, ddsdNN.lPitch)
'                '[UNLOCK]'
'                    BackBufferSurface.Unlock r
'                    'SurfaceXU(2).Unlock r2
'                '[END]'
'                If retVal = -1 Then MsgBox "error!"
'            End If
'[END]'
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 4/22/2006
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function


Function HayUserAbajo(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) _
        And charlist(UserCharIndex).Pos.y <= y
        
End If
End Function

Function PixelPos(ByVal x As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    PixelPos = (TilePixelWidth * x) - TilePixelWidth
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, False, DirGraficos, ClientSetup.byMemory)
          
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128

    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
    'We are done!
    AddtoRichTextBox frmCargando.status, "Hecho.", , , , 1, , False
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer, ByVal engineSpeed As Single, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"

'Set intial user position
UserPos.x = MinXBorder
UserPos.y = MinYBorder

'Fill startup variables
#If ConAlfaB Then
    AlphaTechos = 255
#End If

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


engineBaseSpeed = engineSpeed

'Set FPS value to 100 for startup
FPS = 100
FramesPerSecCounter = 100

'El user empieza caminando.
UserSpeed = 1

ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock


'Set scroll pixels per frame
ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

'Compute offset in pixels when rendering tile buffer.
'We diminish by one to get the top-left corner of the tile for rendering.
TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)


DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
    .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If ClientSetup.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck



Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs


LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

AddtoRichTextBox frmCargando.status, "Cargando Gr�ficos....", 0, 0, 0, , , True
Call LoadGraphics

InitTileEngine = True

End Function

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer)
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    '****** Set main view rectangle ******
    'GetWindowRect DisplayFormhWnd, MainViewRect
    
    With MainViewRect
        .Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
        .Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    If EngineRun Then
    
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame * UserSpeed
                Debug.Print timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame * UserSpeed
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.y = 0
                    UserMoving = False
                End If
            End If
        End If
    
    
        '****** Move screen Left and Right if needed ******
        'If AddtoUserPos.x <> 0 Then
        '    OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.x)))
        '    If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
        '        OffsetCounterX = 0
        '        AddtoUserPos.x = 0
        '        UserMoving = 0
        '    End If
        '****** Move screen Up and Down if needed ******
        'ElseIf AddtoUserPos.y <> 0 Then
        '    OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.y))
        '    If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
        '        OffsetCounterY = 0
        '        AddtoUserPos.y = 0
        '        UserMoving = 0
        '    End If
        'End If


        '****** Update screen ******
        If UserCiego Then
            Call CleanViewPort
        Else
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        End If

        '****** Update screen ******
        'Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)

        'If IScombate Then Call Dialogos.DrawText(260, 260, "MODO COMBATE", vbRed)
        
        Call Dialogos.MostrarTexto
        Call DibujarCartel
        
        Call DialogosClanes.Draw(Dialogos)
        
        Call DrawBackBufferSurface
           
        'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
        While (DirectX.TickCount - fpsLastCheck) \ ClientSetup.FrameInterval < FramesPerSecCounter
            Sleep 5
        Wend
        
        'FPS update
        If fpsLastCheck + 1000 < DirectX.TickCount Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = DirectX.TickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, index As Integer)
ReDim Preserve Grh(1 To index) As Grh
Grh(index).FrameCounter = 1
Grh(index).GrhIndex = GrhIndex
Grh(index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(index).Started = 1
End Sub

Sub CargarAnimsExtra()
    
    'Saao Minimapa.
    Dim DDm As DDSURFACEDESC2
    DDm.lHeight = 101
    DDm.lWidth = 101
    DDm.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
    DDm.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    Set SupMiniMap = DirectDraw.CreateSurface(DDm)
    Set SupBMiniMap = DirectDraw.CreateSurface(DDm)
    
    Call CrearGrh(6580, 1) 'Anim Invent
    Call CrearGrh(534, 2) 'Animacion de teleport
    
End Sub

Function ControlVelocidad(ByVal LastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - LastTime > 20)
End Function


#If ConAlfaB Then

Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
    
    Surface.GetSurfaceDesc ddsdDest
    
    With rRect
        .Left = 0
        .Top = 0
        .Right = ddsdDest.lWidth
        .Bottom = ddsdDest.lHeight
    End With
    
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
    Else
        Modo = 2
    End If
    
    Dim DstLock As Boolean
    DstLock = False
    
    On Local Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
    
    Surface.GetLockedArray dArray()
    
    'If Modo = 1 Then
        If DeNoche = 3 Then
            'Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
                ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
                Modo)
            Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 40, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 0, 0, 60)
        ElseIf DeNoche = 2 Then
            Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 45, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 30, 0, 50)
        ElseIf DeNoche = 0 Then
            Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 40, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 0, 0, 20)
        End If
    'End If
HayErrorAlpha:
    If DstLock = True Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub

#End If

Private Sub CleanViewPort()
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Fills the viewport with black.
'***************************************************
    Dim r As RECT
    Call BackBufferSurface.BltColorFill(r, vbBlack)
End Sub

Private Function GetElapsedTime() As Single
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
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim Color As Long
    
    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffset.x = .MoveOffset.x + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * UserSpeed
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).SpeedCounter > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffset.x >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffset.x <= 0) Then
                    .MoveOffset.x = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffset.y = .MoveOffset.y + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * UserSpeed
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).SpeedCounter > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffset.y >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffset.y <= 0) Then
                    .MoveOffset.y = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            If Not .Arma.WeaponAttack = 1 Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
            Else
                .Arma.WeaponWalk(.Heading).Started = 1
            End If
    
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffset.x
        PixelOffsetY = PixelOffsetY + .MoveOffset.y
        
        If .Head.Head(.Heading).GrhIndex Then
            If Not .invisible Then
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(BackBufferSurface, .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
            
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, .Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                    
                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(BackBufferSurface, .Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                    
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(BackBufferSurface, .Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , , .Arma.WeaponAttack)
                    
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(BackBufferSurface, .Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                
                
                    'Draw name over head
                    If LenB(.Nombre) > 0 Then
                        If Nombres Then  'And Abs(MouseTileX - .Pos.x) < 2 And (Abs(MouseTileY - .Pos.y)) < 2 Then
                            Pos = InStr(.Nombre, "<")
                            If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                            If .priv = 0 Then
                                If .Criminal Then
                                    Color = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                Else
                                    Color = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                End If
                            Else
                                Color = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            End If
                            
                            'Nick
                            line = Left$(.Nombre, Pos - 2)
                            Call Dialogos.DrawText(PixelOffsetX - ((frmMain.TextWidth(line) / 2) - 16), PixelOffsetY + 30, line, Color) 'Call RenderText(PixelOffsetX + 4 - Len(line) * 2, PixelOffsetY + 30, line, color, frmMain.font)
                            
                            'Clan
                            line = mid$(.Nombre, Pos)
                            Call Dialogos.DrawText(PixelOffsetX - ((frmMain.TextWidth(line) / 2) - 16), PixelOffsetY + 45, line, Color) 'Call RenderText(PixelOffsetX + 4 - Len(line) * 2, PixelOffsetY + 30, line, color, frmMain.font)'Call RenderText(PixelOffsetX + 15 - Len(line) * 3, PixelOffsetY + 45, line, color, frmMain.font)
                        End If
                    End If
                End If
            Else
                #If ConAlfaB Then
                    If CharIndex = UserCharIndex Then
                        If .Body.Walk(.Heading).GrhIndex Then _
                            Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                        'Draw Head
                        If .Head.Head(.Heading).GrhIndex Then
                            Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                            
                            'Draw Helmet
                            If .Casco.Head(.Heading).GrhIndex Then _
                                Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                            
                            'Draw Weapon
                            If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                                Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                            
                            'Draw Shield
                            If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                                Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                        
                        
                            'Draw name over head
                            If LenB(.Nombre) > 0 Then
                                If Nombres Then  'And Abs(MouseTileX - .Pos.x) < 2 And (Abs(MouseTileY - .Pos.y)) < 2 Then
                                    Pos = InStr(.Nombre, "<")
                                    If Pos = 0 Then Pos = Len(.Nombre) + 2
                                    
                                    If .priv = 0 Then
                                        If .Criminal Then
                                            Color = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                        Else
                                            Color = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                        End If
                                    Else
                                        Color = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                                    End If
                                    
                                    'Nick
                                    line = Left$(.Nombre, Pos - 2)
                                    Call Dialogos.DrawText(PixelOffsetX - ((frmMain.TextWidth(line) / 2) - 16), PixelOffsetY + 30, line, Color) 'Call RenderText(PixelOffsetX + 4 - Len(line) * 2, PixelOffsetY + 30, line, color, frmMain.font)
                                    
                                    'Clan
                                    line = mid$(.Nombre, Pos)
                                    Call Dialogos.DrawText(PixelOffsetX - ((frmMain.TextWidth(line) / 2) - 16), PixelOffsetY + 45, line, Color) 'Call RenderText(PixelOffsetX + 4 - Len(line) * 2, PixelOffsetY + 30, line, color, frmMain.font)'Call RenderText(PixelOffsetX + 15 - Len(line) * 3, PixelOffsetY + 45, line, color, frmMain.font)
                                End If
                            End If
                        End If
                    End If
                #End If
            End If
        Else 'No head (error condition)
            If Not .invisible Then
            'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(BackBufferSurface, .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
            End If
        End If

        
        'Update dialogs
        If Dialogos.CantidadDialogos > 0 Then
            Call Dialogos.Update_Dialog_Pos(PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, CharIndex)
        End If
        
        'Draw FX
        If .fX <> 0 Then
            #If (ConAlfaB = 1) Then
                If ClientSetup.TransFx = True Then
                    Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .fXGrh, PixelOffsetX + FxData(.fX).OffSetX, PixelOffsetY + FxData(.fX).OffSetY, 1, 1, CharIndex)
                Else
                    Call DDrawTransGrhtoSurface(BackBufferSurface, .fXGrh, PixelOffsetX + FxData(.fX).OffSetX, PixelOffsetY + FxData(.fX).OffSetY, 1, 1, CharIndex)
                End If
            #Else
                Call DDrawTransGrhtoSurface(BackBufferSurface, .fXGrh, PixelOffsetX + FxData(.fX).OffSetX, PixelOffsetY + FxData(.fX).OffSetY, 1, 1, CharIndex)
            #End If
            
            'Check if animation is over
            If .fXGrh.Started = 0 Then _
                .fX = 0
        End If
    End With
End Sub
Public Function SetUserSpeed(ByVal value As Single)
    UserSpeed = value
End Function
