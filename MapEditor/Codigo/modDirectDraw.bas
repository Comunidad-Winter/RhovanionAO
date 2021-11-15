Attribute VB_Name = "modTileEngine"
Option Explicit


Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.X + CX
tY = UserPos.y + CY

End Sub

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
On Error Resume Next

'Update LastChar
If CharIndex > LastChar Then LastChar = CharIndex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(CharIndex).Body = BodyData(Body)
CharList(CharIndex).Head = HeadData(Head)
CharList(CharIndex).Heading = Heading

'Reset moving stats
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.y = 0

'Update position
CharList(CharIndex).Pos.X = X
CharList(CharIndex).Pos.y = y

'Make active
CharList(CharIndex).Active = 1

'Plot on map
MapData(X, y).CharIndex = CharIndex

     Call DibujarMiniMapa

End Sub







Sub EraseChar(CharIndex As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
If CharIndex = 0 Then Exit Sub
'Make un-active
CharList(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.y).CharIndex = 0

'Update NumChars
NumChars = NumChars - 1

     Call DibujarMiniMapa

End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).Pos.X
y = CharList(CharIndex).Pos.y

'Figure out which way to move
Select Case nHeading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.y = nY
MapData(X, y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
Dim X As Integer
Dim y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

X = CharList(CharIndex).Pos.X
y = CharList(CharIndex).Pos.y

addX = nX - X
addY = nY - y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.y = nY
MapData(X, y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

     Call DibujarMiniMapa

End Sub


Function NextOpenChar() As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function

Function LegalPos(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

LegalPos = True

'Check to see if its out of bounds
If X - 8 < 1 Or X + 8 > 100 Or y - 6 < 1 Or y + 6 > 100 Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, y).CharIndex > 0 Then
    LegalPos = False
    Exit Function
End If

End Function




Function InMapLegalBounds(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < XMinMapSize Or X > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

' [Loopzer]
Public Sub DePegar()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
             MapData(X + DeSeleccionOX, y + DeSeleccionOY) = DeSeleccionMap(X, y)
        Next
    Next
End Sub
Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SobreX, y + SobreY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             MapData(X + SobreX, y + SobreY) = SeleccionMap(X, y)
        Next
    Next
    Seleccionando = False
End Sub
Public Sub AccionSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + y
        Next
    Next
    Seleccionando = False
End Sub

Public Sub BlockearSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer
    Dim Vacio As MapBlock
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             If MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 1 Then
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 0
             Else
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 1
            End If
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CortarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    CopiarSeleccion
    Dim X As Integer
    Dim y As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             MapData(X + SeleccionIX, y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CopiarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim y As Integer
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
End Sub
Public Sub GenerarVista()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
   ' hacer una llamada a un seter o geter , es mas lento q una variable
   ' con esto hacemos q no este preguntando a el objeto cadavez
   ' q dibuja , Render mas rapido ;)
    VerBlockeados = frmMain.cVerBloqueos.value
    VerTriggers = frmMain.cVerTriggers.value
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    
End Sub
' [/Loopzer]
Public Sub RenderScreen(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 31/05/06 by GS
'Last modified: 21/11/07 By Loopzer
'Last modifier: 24/11/08 by GS
'*************************************************

On Error Resume Next
Dim y       As Integer              'Keeps track of where on map we are
Dim X       As Integer
Dim MinY    As Integer              'Start Y pos on current map
Dim MaxY    As Integer              'End Y pos on current map
Dim MinX    As Integer              'Start X pos on current map
Dim MaxX    As Integer              'End X pos on current map
Dim ScreenX As Integer              'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim R       As RECT
Dim Sobre   As Integer
Dim Moved   As Byte
Dim iPPx    As Integer              'Usado en el Layer de Chars
Dim iPPy    As Integer              'Usado en el Layer de Chars
Dim grh     As grh                  'Temp Grh for show tile and blocked
Dim bCapa    As Byte                 'cCapas ' 31/05/2006 - GS, control de Capas
Dim SelRect As RECT
Dim rSourceRect         As RECT     'Usado en el Layer 1
Dim iGrhIndex           As Integer  'Usado en el Layer 1
Dim PixelOffsetXTemp    As Integer  'For centering grhs
Dim PixelOffsetYTemp    As Integer
Dim TempChar            As Char

Dim colorlist(3) As Long

colorlist(0) = D3DColorXRGB(255, 200, 0)
colorlist(1) = D3DColorXRGB(255, 200, 0)
colorlist(2) = D3DColorXRGB(255, 200, 0)
colorlist(3) = D3DColorXRGB(255, 200, 0)

Map_LightsRender

MinY = (TileY - (WindowTileHeight \ 2)) - TileBufferSize
MaxY = (TileY + (WindowTileHeight \ 2)) + TileBufferSize
MinX = (TileX - (WindowTileWidth \ 2)) - TileBufferSize
MaxX = (TileX + (WindowTileWidth \ 2)) + TileBufferSize
' 31/05/2006 - GS, control de Capas
If Val(frmMain.cCapas.Text) >= 1 And (frmMain.cCapas.Text) <= 4 Then
    bCapa = Val(frmMain.cCapas.Text)
Else
    bCapa = 1
End If
GenerarVista 'Loopzer
ScreenY = -8
For y = (MinY) To (MaxY)
    ScreenX = -8
    For X = (MinX) To (MaxX)
        If InMapBounds(X, y) Then
            If X > 100 Or y < 1 Then Exit For ' 30/05/2006

            'Layer 1 **********************************
            If SobreX = X And SobreY = y Then
                ' Pone Grh !
                Sobre = -1
                If frmMain.cSeleccionarSuperficie.value = True Then
                    Sobre = MapData(X, y).Graphic(bCapa).Grh_Index
                    If frmConfigSup.MOSAICO.value = vbChecked Then
                        Dim aux As Integer
                        Dim dy As Integer
                        Dim dX As Integer
                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo.Text)
                            dX = Val(frmConfigSup.DMAncho.Text)
                        Else
                            dy = 0
                            dX = 0
                        End If
                        If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                            aux = Val(frmMain.cGrh.Text) + _
                            (((y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
                            If MapData(X, y).Graphic(bCapa).Grh_Index <> aux Then
                                MapData(X, y).Graphic(bCapa).Grh_Index = aux
                                Grh_Initialize MapData(X, y).Graphic(bCapa), aux
                            End If
                        Else
                            aux = Val(frmMain.cGrh.Text) + _
                            (((y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
                            If MapData(X, y).Graphic(bCapa).Grh_Index <> aux Then
                                MapData(X, y).Graphic(bCapa).Grh_Index = aux
                                Grh_Initialize MapData(X, y).Graphic(bCapa), aux
                            End If
                        End If
                    Else
                        If MapData(X, y).Graphic(bCapa).Grh_Index <> Val(frmMain.cGrh.Text) Then
                            MapData(X, y).Graphic(bCapa).Grh_Index = Val(frmMain.cGrh.Text)
                            Grh_Initialize MapData(X, y).Graphic(bCapa), Val(frmMain.cGrh.Text)
                        End If
                    End If
                End If
            Else
                Sobre = -1
            End If
            If VerCapa1 Then
                With MapData(X, y).Graphic(1)
                    modGrh.Grh_Render MapData(X, y).Graphic(1), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value
                End With
            End If
            'Layer 2 **********************************
            If MapData(X, y).Graphic(2).Grh_Index <> 0 And VerCapa2 Then
                modGrh.Grh_Render MapData(X, y).Graphic(2), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
            End If
            
            If Sobre >= 0 Then
                If MapData(X, y).Graphic(bCapa).Grh_Index <> Sobre Then
                MapData(X, y).Graphic(bCapa).Grh_Index = Sobre
                Grh_Initialize MapData(X, y).Graphic(bCapa), Sobre
                End If
            End If
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If y > 100 Then Exit For
Next y
ScreenY = -8
For y = (MinY) To (MaxY)   '- 8+ 8
    ScreenX = -8
    For X = (MinX) To (MaxX)   '- 8 + 8
        If InMapBounds(X, y) Then
            If X > 100 Or X < -3 Then Exit For ' 30/05/2006

            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
             'Object Layer **********************************
             If MapData(X, y).OBJInfo.objindex <> 0 And VerObjetos Then
                modGrh.Grh_Render MapData(X, y).ObjGrh, iPPx, iPPy, MapData(X, y).light_value, True
             End If
            
                  'Char layer **********************************
                 If MapData(X, y).CharIndex <> 0 And VerNpcs Then
                 
                     TempChar = CharList(MapData(X, y).CharIndex)
                 
                     PixelOffsetXTemp = PixelOffsetX
                     PixelOffsetYTemp = PixelOffsetY
                    
                   'Dibuja solamente players
                   If TempChar.Head.Head(TempChar.Heading).Grh_Index <> 0 Then
                     'Draw Body
                     modGrh.Grh_Render TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True
                     'Draw Head
                     modGrh.Grh_Render TempChar.Head.Head(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True
                   Else: modGrh.Grh_Render TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True
                   End If
                 End If
             'Layer 3 *****************************************
             If MapData(X, y).Graphic(3).Grh_Index <> 0 And VerCapa3 Then
                'Draw
                modGrh.Grh_Render MapData(X, y).Graphic(3), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
             End If
             
             If MapData(X, y).particle_group_index Then
                modDXEngine.DXEngine_ParticleGroupRender MapData(X, y).particle_group_index, iPPx, iPPy
             End If
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next y
'Tiles blokeadas, techos, triggers , seleccion
ScreenY = -8
For y = (MinY) To (MaxY)
    ScreenX = -8
    For X = (MinX) To (MaxX)
        If X < 101 And X > 0 And y < 101 And y > 0 Then ' 30/05/2006
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
            If MapData(X, y).Graphic(4).Grh_Index <> 0 _
            And (frmMain.mnuVerCapa4.Checked = True) Then
                'Draw
                modGrh.Grh_Render MapData(X, y).Graphic(4), iPPx, iPPy, MapData(X, y).light_value, True
            End If
            If MapData(X, y).TileExit.Map <> 0 And VerTranslados Then
                grh.Grh_Index = 3
                grh.frame_counter = 1
                grh.Started = 0
                modGrh.Grh_Render grh, iPPx, iPPy, MapData(X, y).light_value, True
            End If
            
            If MapData(X, y).light_index Then
                grh.Grh_Index = 4
                grh.frame_counter = 1
                grh.Started = 0
                modGrh.Grh_Render grh, iPPx, iPPy, colorlist, True
            End If
            
            'Show blocked tiles
            If VerBlockeados And MapData(X, y).Blocked = 1 Then
                grh.Grh_Index = 4
                grh.frame_counter = 1
                grh.Started = 0
                modGrh.Grh_Render grh, iPPx, iPPy, MapData(X, y).light_value, True
            End If
            If VerGrilla Then
                'Grilla 24/11/2008 by GS
                modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 32, RGB(255, 255, 255)
                modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 1, RGB(255, 255, 255)
            End If
            If VerTriggers Then
                Call DrawText(PixelPos(ScreenX), PixelPos(ScreenY), Str(MapData(X, y).Trigger), vbRed)
            End If
            If Seleccionando Then
                'If ScreenX >= SeleccionIX And ScreenX <= SeleccionFX And ScreenY >= SeleccionIY And ScreenY <= SeleccionFY Then
                    If X >= SeleccionIX And y >= SeleccionIY Then
                        If X <= SeleccionFX And y <= SeleccionFY Then
                            modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 32, RGB(100, 255, 255)
                        End If
                    End If
            End If

        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next y

End Sub



Public Sub DrawText(lngXPos As Integer, lngYPos As Integer, strText As String, lngColor As Long)
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
    If LenB(strText) <> 0 Then
        Call modDXEngine.DXEngine_TextRender(1, strText, lngXPos, lngYPos, D3DColorXRGB(255, 255, 255))
    End If
End Sub

Function PixelPos(X As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 15/10/06 by GS
'*************************************************
    'Fill startup variables
    DisplayFormhWnd = setDisplayFormhWnd
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    '[GS] 02/10/2006
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
    
    MainViewWidth = (TilePixelWidth * WindowTileWidth)
    MainViewHeight = (TilePixelHeight * WindowTileHeight)
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    InitTileEngine = True
    EngineRun = True
    DoEvents
End Function

Public Sub LightSet(ByVal X As Byte, ByVal y As Byte, ByVal Rounded As Boolean, ByVal Range As Integer, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim ix As Integer
    Dim iy As Integer
    Dim i As Integer
    
    If Rounded Then
        For i = 1 To Light_Count
            If Light_Count = 0 Then Exit For
            If Lights(i).Active = 0 Then
                Exit For
            End If
        Next i
        If i > Light_Count Then
            Light_Count = Light_Count + 1
            i = Light_Count
        End If
        MapData(X, y).light_index = i
        ReDim Preserve Lights(1 To Light_Count) As Light
        Lights(i).Active = True
        Lights(i).map_x = X
        Lights(i).map_y = y
        Lights(i).X = X * 32
        Lights(i).y = y * 32
        Lights(i).Range = Range
        Lights(i).RGBCOLOR.A = 255
        Lights(i).RGBCOLOR.R = R
        Lights(i).RGBCOLOR.G = G
        Lights(i).RGBCOLOR.B = B
    Else
        'Set up light borders
        min_x = X - Range
        min_y = y - Range
        max_x = X + Range
        max_y = y + Range
    
        If InMapBounds(min_x, min_y) Then
            MapData(min_x, min_y).base_light(2) = True
            MapData(min_x, min_y).light_base_value(2) = D3DColorXRGB(R, G, B)
        End If
        If InMapBounds(min_x, max_y) Then
            MapData(min_x, max_y).base_light(3) = True
            MapData(min_x, max_y).light_base_value(3) = D3DColorXRGB(R, G, B)
        End If
        If InMapBounds(max_x, min_y) Then
            MapData(max_x, min_y).base_light(0) = True
            MapData(max_x, min_y).light_base_value(0) = D3DColorXRGB(R, G, B)
        End If
        If InMapBounds(max_x, max_y) Then
            MapData(max_x, max_y).base_light(1) = True
            MapData(max_x, max_y).light_base_value(1) = D3DColorXRGB(R, G, B)
        End If
        
        'Upper Border
        For ix = min_x + 1 To max_x - 1
            If InMapBounds(ix, min_y) Then
                MapData(ix, min_y).base_light(0) = True
                MapData(ix, min_y).light_base_value(0) = D3DColorXRGB(R, G, B)
                MapData(ix, min_y).base_light(2) = True
                MapData(ix, min_y).light_base_value(2) = D3DColorXRGB(R, G, B)
            End If
        Next ix
        
        'Lower Border
        For ix = min_x + 1 To max_x - 1
            If InMapBounds(ix, max_y) Then
                MapData(ix, max_y).base_light(3) = True
                MapData(ix, max_y).light_base_value(3) = D3DColorXRGB(R, G, B)
                MapData(ix, max_y).base_light(1) = True
                MapData(ix, max_y).light_base_value(1) = D3DColorXRGB(R, G, B)
            End If
        Next ix
        
        'Right Border
        For iy = min_y + 1 To max_y - 1
            If InMapBounds(max_x, iy) Then
                MapData(max_x, iy).base_light(1) = True
                MapData(max_x, iy).light_base_value(1) = D3DColorXRGB(R, G, B)
                MapData(max_x, iy).base_light(0) = True
                MapData(max_x, iy).light_base_value(0) = D3DColorXRGB(R, G, B)
            End If
        Next iy
        
        'Left Border
        For iy = min_y + 1 To max_y - 1
            If InMapBounds(min_x, iy) Then
                MapData(min_x, iy).base_light(3) = True
                MapData(min_x, iy).light_base_value(3) = D3DColorXRGB(R, G, B)
                MapData(min_x, iy).base_light(2) = True
                MapData(min_x, iy).light_base_value(2) = D3DColorXRGB(R, G, B)
            End If
        Next iy
        
        'Left Border
        For iy = min_y + 1 To max_y - 1
            For ix = min_x + 1 To max_x - 1
                If InMapBounds(ix, iy) Then
                    MapData(ix, iy).base_light(3) = True
                    MapData(ix, iy).light_base_value(3) = D3DColorXRGB(R, G, B)
                    MapData(ix, iy).base_light(2) = True
                    MapData(ix, iy).light_base_value(2) = D3DColorXRGB(R, G, B)
                    MapData(ix, iy).base_light(1) = True
                    MapData(ix, iy).light_base_value(1) = D3DColorXRGB(R, G, B)
                    MapData(ix, iy).base_light(0) = True
                    MapData(ix, iy).light_base_value(0) = D3DColorXRGB(R, G, B)
                End If
            Next ix
        Next iy
    End If
End Sub


Public Sub Map_LightsRender()
    Dim i As Integer
    
    Call Map_LightsClear
    
    For i = 1 To Light_Count
        Map_LightRender (i)
    Next i
End Sub

Private Function Map_LightsClear()
    Dim X As Integer
    Dim y As Integer
    
    Dim AmbientColor As D3DCOLORVALUE
    Dim Color As Long
    
    Meteo.Get_AmbientLight AmbientColor
    Color = D3DColorXRGB(AmbientColor.R, AmbientColor.G, AmbientColor.B)
    
    For X = 1 To 100
        For y = 1 To 100
            If InMapBounds(X, y) Then
                With MapData(X, y)
                    If .base_light(0) Then 'Si tiene luz propia, la seteamos.
                        .light_value(0) = .light_base_value(0)
                    Else
                        .light_value(0) = Color
                    End If
                    If .base_light(1) Then
                        .light_value(1) = .light_base_value(1)
                    Else
                        .light_value(1) = Color
                    End If
                    If .base_light(2) Then
                        .light_value(2) = .light_base_value(2)
                    Else
                        .light_value(2) = Color
                    End If
                    If .base_light(3) Then
                        .light_value(3) = .light_base_value(3)
                    Else
                        .light_value(3) = Color
                    End If
                End With
            End If
        Next y
    Next X
End Function

Private Sub Map_LightRender(ByVal light_index As Integer)
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Color As Long
    Dim Ya As Integer
    Dim Xa As Integer
    
    Dim TileLight As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
    
    Dim XCoord As Integer
    Dim YCoord As Integer
        
        LightColor = Lights(light_index).RGBCOLOR
        Meteo.Get_AmbientLight AmbientColor
        
        If Not Lights(light_index).Active = True Then Exit Sub
        
        min_x = Lights(light_index).map_x - Lights(light_index).Range
        max_x = Lights(light_index).map_x + Lights(light_index).Range
        min_y = Lights(light_index).map_y - Lights(light_index).Range
        max_y = Lights(light_index).map_y + Lights(light_index).Range
        
        For Ya = min_y To max_y
            For Xa = min_x To max_x
                If InMapBounds(Xa, Ya) Then
                    XCoord = Xa * 32
                    YCoord = Ya * 32
                    'Color = LightCalculate(lights(light_index).range, lights(light_index).x, lights(light_index).y, XCoord, YCoord, mapdata(Xa, Ya).light_value(1), LightColor, AmbientColor)
                    MapData(Xa, Ya).light_value(1) = LightCalculate(Lights(light_index).Range, Lights(light_index).X, Lights(light_index).y, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)

                    XCoord = Xa * 32 + 32
                    YCoord = Ya * 32
                    MapData(Xa, Ya).light_value(3) = LightCalculate(Lights(light_index).Range, Lights(light_index).X, Lights(light_index).y, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                    XCoord = Xa * 32
                    YCoord = Ya * 32 + 32
                    MapData(Xa, Ya).light_value(0) = LightCalculate(Lights(light_index).Range, Lights(light_index).X, Lights(light_index).y, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
    
                    XCoord = Xa * 32 + 32
                    YCoord = Ya * 32 + 32
                    MapData(Xa, Ya).light_value(2) = LightCalculate(Lights(light_index).Range, Lights(light_index).X, Lights(light_index).y, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
                End If
            Next Xa
        Next Ya
End Sub

Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
    
    Dim CurrentColor As D3DCOLORVALUE
    
    pRadio = cRadio * 32
    
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
    
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
    
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(CurrentColor.R, CurrentColor.G, CurrentColor.B)
        If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function

Public Sub LightDestroy(ByVal X As Byte, ByVal y As Byte)
    Dim temp As Light
    Dim i As Byte
    If MapData(X, y).light_index Then
        Lights(MapData(X, y).light_index).Active = False
        MapData(X, y).light_index = 0
    Else
        MapData(X, y).base_light(0) = False
        MapData(X, y).base_light(1) = False
        MapData(X, y).base_light(2) = False
        MapData(X, y).base_light(3) = False
    End If
End Sub

Public Sub LightDestroyAll()
    Dim X As Integer
    Dim y As Integer
    For X = 1 To 100
        For y = 1 To 100
        Call LightDestroy(X, y)
        Next y
    Next X
End Sub

'MINIMAPA
Public Sub DibujarMiniMapa()
Dim map_x As Long, map_y As Long
 
    For map_y = 1 To 100
        For map_x = 1 To 100
           If MapData(map_x, map_y).Graphic(1).Grh_Index > 0 Then



            If MapData(map_x, map_y).Graphic(1).Grh_Index > 0 Then
                SetPixel frmMain.Minimap.hdc, map_x, map_y, grh_list(MapData(map_x, map_y).Graphic(1).Grh_Index).MiniMap_color
            End If


            If MapData(map_x, map_y).Graphic(2).Grh_Index > 0 Then
                SetPixel frmMain.Minimap.hdc, map_x, map_y, grh_list(MapData(map_x, map_y).Graphic(2).Grh_Index).MiniMap_color
            End If
        
            
            End If
        Next map_x
    Next map_y
   
    SetPixel frmMain.Minimap.hdc, UserPos.X, UserPos.y, RGB(255, 0, 0)
    SetPixel frmMain.Minimap.hdc, UserPos.X + 1, UserPos.y, RGB(255, 0, 0)
    SetPixel frmMain.Minimap.hdc, UserPos.X - 1, UserPos.y, RGB(255, 0, 0)
    SetPixel frmMain.Minimap.hdc, UserPos.X, UserPos.y - 1, RGB(255, 0, 0)
    SetPixel frmMain.Minimap.hdc, UserPos.X, UserPos.y + 1, RGB(255, 0, 0)
 
  Dim MinX As Byte
    Dim MinY As Byte
    Dim MaxX As Byte
    Dim MaxY As Byte
   
    map_x = 0
    map_y = 0
   
    MinX = UserPos.X - 5
    MaxX = UserPos.X + 5
   
    MinY = UserPos.y - 5
    MaxY = UserPos.y + 5
   
    For map_y = MinY To MaxY
        SetPixel frmMain.Minimap.hdc, MinX, map_y, RGB(255, 255, 255)
    Next map_y
   
    For map_y = MinY To MaxY
        SetPixel frmMain.Minimap.hdc, MaxX, map_y, RGB(255, 255, 255)
    Next map_y
   
    For map_x = MinX To MaxX
        SetPixel frmMain.Minimap.hdc, map_x, MinY, RGB(255, 255, 255)
    Next map_x
 
    For map_x = MinX To MaxX
        SetPixel frmMain.Minimap.hdc, map_x, MaxY, RGB(255, 255, 255)
    Next map_x
 
    frmMain.Minimap.Refresh
 
End Sub
'MINIMAPA
