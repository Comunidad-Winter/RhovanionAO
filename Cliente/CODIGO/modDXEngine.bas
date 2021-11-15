Attribute VB_Name = "modDXEngine"
Option Explicit

Private Const DegreeToRadian As Single = 0.0174532925

'***************************
'Estructures
'***************************
'This structure describes a transformed and lit vertex.
Private Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    Color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Private Type tGraphicChar
    Src_X As Integer
    Src_Y As Integer
End Type

Private Type tGraphicFont
    texture_index As Long
    Caracteres(0 To 255) As tGraphicChar 'Ascii Chars
    Char_Size As Byte 'In pixels
End Type

Private Type DXFont
    dFont As D3DXFont
    Size As Integer
End Type

Private Type tParticle
    x As Single
    y As Single
    vX As Single
    vY As Single
    
    Screen_X As Single
    Screen_Y As Single
    
    Moved_X As Single
    Moved_Y As Single
    
    Created As Long
    Alive As Byte
    CurrentColor(1 To 4) As D3DCOLORVALUE
    rgb_list(3) As Long
    
    angle As Single
    
    texture_index As Integer
    
    Particle_LifeTime As Long
    Used As Boolean
    
    Delay As Integer
    DelayCounter As Integer
End Type

Private Type tParticle_Emisor
    x1 As Integer
    x2 As Integer
    Y1 As Integer
    Y2 As Integer
    
    vX1 As Integer
    vY1 As Integer
    vX2 As Integer
    vY2 As Integer
    
    MoveX As Boolean
    MoveY As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    
    Particle_count As Integer
    
    Particle_Speed As Single
    Particle_Frame_Counter As Single
    
    Particle() As tParticle
    
    Gravity As Single
    
    Friction As Single
    
    RatioFriction As Single
    ColorVariation As Byte
    
    StartColor(1 To 4) As D3DCOLORVALUE
    EndColor(1 To 4) As D3DCOLORVALUE
    
    PLt1 As Integer
    PLt2 As Integer
    
    alpha_blend As Boolean
    
    WindDirection As Byte
    Wind As Single
    
    Spin As Byte
    SpinH As Integer
    SpinL As Integer
    
    Bounce_Strength As Integer
    Bounce_Y As Integer
    
    texture_count As Integer
    texture_index() As Integer
    texture_size() As Integer
    
    Ratio As Integer
    RatioVariation As Integer
    CurrentRatio As Single
    
    ParticleGroup_Type As Integer
    
    ParticlesLeft As Integer
    
    StartParticlesDestroy As Boolean 'Time
    
    KillWhenAtTarget As Boolean
    Target_X As Integer
    Target_Y As Integer
    
    StopAtTargetRatio As Boolean 'If False then it loops
    TargetRatio As Integer
    
    KillType As Byte
    
    Delay1 As Integer
    Delay2 As Integer
End Type

Private Type tParticle_Group
    Active As Byte
    'Pointers to tileengine. NO ME GUSTA ESTO MEJORAR ALGUN DIA :P
    map_x As Integer
    map_y As Integer
    char_index As Integer
    ParticleGroup_Lifetime As Long
    Created As Long
    ParticleEmisor As tParticle_Emisor
    ParticleEmisor_Count As Integer
End Type

Public Enum e_ParticleType
    Rain = 7
End Enum

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

'***************************
'Variables
'***************************
'Major DX Objects
Public dx As DirectX8
Public d3d As Direct3D8
Public ddevice As Direct3DDevice8
Public d3dx As D3DX8

Dim d3dpp As D3DPRESENT_PARAMETERS

'Texture Manager for Dinamic Textures
Dim DXPool As New clsTextureManager

'Main form handle
Dim form_hwnd As Long

'Display variables
Dim screen_hwnd As Long
Dim screen_width As Long
Dim screen_height As Long

'FPS Counters
Dim fps_last_time As Long 'When did we last check the frame rate?
Dim fps_frame_counter As Long 'How many frames have been drawn
Dim fps As Long 'What the current frame rate is.....

Dim particle_timer As Single

Dim engine_render_started As Boolean

'Graphic Font List
Dim gfont_list() As tGraphicFont
Dim gfont_count As Long
Dim gfont_last As Long

'Font List
Private font_list() As DXFont
Private font_count As Integer

'Particles
Private particle_group_list() As tParticle_Group
Private particle_group_count As Integer

Private particle_group_data_count As Integer
Private particle_group_data() As tParticle_Group

'***************************
'Constants
'***************************
'Engine
Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
'PI
Private Const PI As Single = 3.14159265358979

'Old fashion BitBlt functions
Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const SRCAND = &H8800C6
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcsrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Initialization
Public Function DXEngine_Initialize(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean)
'On Error GoTo errhandler
    Dim d3dcaps As D3DCAPS8
    Dim d3ddm As D3DDISPLAYMODE
    
    DXEngine_Initialize = True
    
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    '*******************************
    'Initialize root DirectX8 objects
    '*******************************
    Set dx = New DirectX8
    'Create the Direct3D object
    Set d3d = dx.Direct3DCreate
    'Create helper class
    Set d3dx = New D3DX8
    
    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    DevType = D3DDEVTYPE_HAL
    'Get the capabilities of the Direct3D device that we specify. In this case,
    'we'll be using the adapter default (the primiary card on the system).
    Call d3d.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, d3dcaps)
    'Grab some information about the current display mode.
    Call d3d.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, d3ddm)
    
    'Now we'll go ahead and fill the D3DPRESENT_PARAMETERS type.
    With d3dpp
        .windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = d3ddm.Format 'current display depth
    End With
    'create device
    Set ddevice = d3d.CreateDevice(D3DADAPTER_DEFAULT, DevType, screen_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)

    DeviceRenderStates
    
    '****************************************************
    'Inicializamos el manager de texturas
    '****************************************************
    Call DXPool.Texture_Initialize(500)
    
    '****************************************************
    'Clears the buffer to start rendering
    '****************************************************
    Device_Clear
    '****************************************************
    'Load Misc
    '****************************************************
    LoadGraphicFonts
    LoadFonts
    LoadParticles
    Particles_Initialize
    
    Exit Function
ErrHandler:
    DXEngine_Initialize = False
End Function

Public Function DXEngine_BeginRender() As Boolean
On Error GoTo ErrorHandler:
    DXEngine_BeginRender = True
    
    'Check if we have the device
    If ddevice.TestCooperativeLevel <> D3D_OK Then
        Do
            DoEvents
        Loop While ddevice.TestCooperativeLevel = D3DERR_DEVICELOST
        
        DXPool.Texture_Remove_All
        Fonts_Destroy
        Device_Reset
        
        DeviceRenderStates
        LoadFonts
        LoadGraphicFonts
    End If
    
    '****************************************************
    'Render
    '****************************************************
    '*******************************
    'Erase the backbuffer so that it can be drawn on again
    Device_Clear
    '*******************************
    '*******************************
    'Start the scene
    ddevice.BeginScene
    '*******************************
    
    engine_render_started = True
Exit Function
ErrorHandler:
    DXEngine_BeginRender = False
    MsgBox "Error in Engine_Render_Start: " & Err.Number & ": " & Err.Description
End Function

Public Function DXEngine_EndRender() As Boolean
On Error GoTo ErrorHandler:
    DXEngine_EndRender = True

    If engine_render_started = False Then
        Exit Function
    End If
    
    '*******************************
    'End scene
    ddevice.EndScene
    '*******************************
    
    '*******************************
    'Flip the backbuffer to the screen
    Device_Flip
    '*******************************
    
    '*******************************
    'Calculate current frames per second
    If GetTickCount >= (fps_last_time + 1000) Then
        fps = fps_frame_counter
        fps_frame_counter = 0
        fps_last_time = GetTickCount
    Else
        fps_frame_counter = fps_frame_counter + 1
    End If
    '*******************************
    

    
    
    engine_render_started = False
Exit Function
ErrorHandler:
    DXEngine_EndRender = False
    MsgBox "Error in Engine_Render_End: " & Err.Number & ": " & Err.Description
End Function

Private Sub Device_Clear()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1#, 0
End Sub

Private Function Device_Reset() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Resets the device
'**************************************************************
On Error GoTo ErrHandler:
'On Error Resume Next

    'Be sure the scene is finished
    ddevice.EndScene
    'Reset device
    ddevice.Reset d3dpp
    
    DeviceRenderStates
       
Exit Function
ErrHandler:
    Device_Reset = Err.Number
End Function
Public Sub DXEngine_TextureRenderAdvance(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal Src_X As Long, ByVal Src_Y As Long, _
                                             ByVal dest_width As Long, ByVal dest_height As Long, ByVal src_width As Long, ByVal src_height As Long, ByRef rgb_list() As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'This sub allow texture resizing
'
'**************************************************************

    
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim Texture As Direct3DTexture8
    Dim texture_width As Integer
    Dim texture_height As Integer

    'rgb_list(0) = RGB(255, 255, 255)
    'rgb_list(1) = RGB(255, 255, 255)
    'rgb_list(2) = RGB(255, 255, 255)
    'rgb_list(3) = RGB(255, 255, 255)
    
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    With src_rect
        .bottom = Src_Y + src_height
        .Right = Src_X + src_width
        .top = Src_Y
        .left = Src_X
    End With
    
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, rgb_list(), texture_width, texture_height, angle
    
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub
Public Sub DXEngine_TextureRender(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByRef rgb_list() As Long, ByVal Src_X As Long, _
                                            ByVal Src_Y As Long, ByVal dest_width As Long, ByVal dest_height As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'This sub doesnt allow texture resizing
'
'**************************************************************
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim texture_height As Integer
    Dim texture_width As Integer
    Dim Texture As Direct3DTexture8
    
    'Set up the source rectangle
    With src_rect
        .bottom = Src_Y + src_height - 1
        .left = Src_X
        .Right = Src_X + src_width - 1
        .top = Src_Y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    'ESTO NO ME GUSTA
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, rgb_list(), texture_height, texture_width, angle
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    'Enable alpha-blending
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    'Turn off alphablending after we're done
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub
Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.specular = specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef texture_width As Integer, Optional ByRef texture_height As Integer, Optional ByVal angle As Single)
'**************************************************************
'Authors: Aaron Perkins;
'Last Modify Date: 5/07/2002
'
' * v1 *    v3
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' * v0 *    v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left - 1) / 2
        y_center = dest.top + (dest.bottom - dest.top - 1) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If texture_width And texture_height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / texture_width, (src.bottom) / texture_height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If texture_width And texture_height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / texture_width, src.top / texture_height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If texture_width And texture_height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right) / texture_width, (src.bottom) / texture_height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If texture_width And texture_height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right) / texture_width, src.top / texture_height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 0)
    End If
End Sub

Public Sub DXEngine_GraphicTextRender(Font_Index As Integer, ByVal Text As String, ByVal top As Long, ByVal left As Long, _
                                  ByVal Color As Long)

    If Len(Text) > 255 Then Exit Sub
    
    Dim i As Byte
    Dim x As Integer
    Dim y As Integer
    Dim rgb_list(3) As Long
    
    For i = 0 To 3
        rgb_list(i) = Color
    Next i
    
    x = -1
    Dim Char As Integer
    For i = 1 To Len(Text)
        Char = AscB(Mid$(Text, i, 1)) - 32
        
        If Char = 0 Then
            x = x + 1
        Else
            x = x + 1
            Call DXEngine_TextureRenderAdvance(gfont_list(Font_Index).texture_index, left + x * gfont_list(Font_Index).Char_Size, _
                                                        top, gfont_list(Font_Index).Caracteres(Char).Src_X, gfont_list(Font_Index).Caracteres(Char).Src_Y, _
                                                            gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, _
                                                                rgb_list(), False)
        End If
    Next i
    
    
    
End Sub

Public Sub DXEngine_Deinitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next

    'El manager de texturas es ahora independiente del engine.
    Call DXPool.Texture_Remove_All
    
    Set d3dx = Nothing
    Set ddevice = Nothing
    Set d3d = Nothing
    Set dx = Nothing
    Set DXPool = Nothing
End Sub

Private Sub LoadChars(ByVal Font_Index As Integer)
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    
    For i = 0 To 255
        With gfont_list(Font_Index).Caracteres(i)
            x = (i Mod 16) * gfont_list(Font_Index).Char_Size
            If x = 0 Then '16 chars per line
                y = y + 1
            End If
            .Src_X = x
            .Src_Y = (y * gfont_list(Font_Index).Char_Size) - gfont_list(Font_Index).Char_Size
        End With
    Next i
End Sub
Public Sub LoadGraphicFonts()
    Dim i As Byte
    Dim file_path As String

    file_path = resource_path & PATH_INIT & "\GUIFonts.ini"

    If General_File_Exists(file_path, vbArchive) Then
        gfont_count = General_Var_Get(file_path, "INIT", "FontCount")
        If gfont_count > 0 Then
            ReDim gfont_list(1 To gfont_count) As tGraphicFont
            For i = 1 To gfont_count
                With gfont_list(i)
                    .Char_Size = General_Var_Get(file_path, "FONT" & i, "Size")
                    .texture_index = General_Var_Get(file_path, "FONT" & i, "Graphic")
                    If .texture_index > 0 Then Call DXPool.Texture_Load(.texture_index, 0)
                    LoadChars (i)
                End With
            Next i
        End If
    End If
End Sub

Public Sub DXEngine_StatsRender()
    'fps
    Call DXEngine_TextRender(1, fps & " FPS", 0, 0, D3DColorXRGB(255, 255, 255))
End Sub

Private Sub Device_Flip()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Draw the graphics to the front buffer.
    ddevice.Present ByVal 0&, ByVal 0&, screen_hwnd, ByVal 0&
End Sub

Private Sub DeviceRenderStates()
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        .SetVertexShader FVF
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        'No se para q mierda sera esto.
        '.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        '.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        '.SetRenderState D3DRS_ZENABLE, True
        '.SetRenderState D3DRS_ZWRITEENABLE, False
        
        'Particle engine settings
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        '.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        '.SetRenderState D3DRS_POINTSCALE_ENABLE, 0


    End With
End Sub

Private Sub Font_Make(ByVal Style As String, ByVal Size As Long, ByVal italic As Boolean, ByVal bold As Boolean)
    font_count = font_count + 1
    ReDim Preserve font_list(1 To font_count)
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.Name = Style
    fnt.Size = Size
    fnt.bold = bold
    fnt.italic = italic
    Set font_desc = fnt
    font_list(font_count).Size = Size
    Set font_list(font_count).dFont = d3dx.CreateFont(ddevice, font_desc.hFont)
End Sub

Private Sub LoadFonts()
    Dim num_fonts As Integer
    Dim i As Integer
    Dim file_path As String
    
    file_path = resource_path & PATH_INIT & "\fonts.ini"
    
    If Not General_File_Exists(file_path, vbArchive) Then Exit Sub
    
    num_fonts = General_Var_Get(file_path, "INIT", "FontCount")
    
    For i = 1 To num_fonts
        Call Font_Make(General_Var_Get(file_path, "FONT" & i, "Name"), General_Var_Get(file_path, "FONT" & i, "Size"), General_Var_Get(file_path, "FONT" & i, "Cursiva"), General_Var_Get(file_path, "FONT" & i, "Negrita"))
    Next i
End Sub
Public Sub DXEngine_TextRender(ByVal Font_Index As Integer, ByVal Text As String, ByVal left As Integer, ByVal top As Integer, ByVal Color As Long, Optional ByVal Alingment As Byte = DT_LEFT, Optional ByVal width As Integer = 0, Optional ByVal height As Integer = 0)
    If Not Font_Check(Font_Index) Then Exit Sub
    
    Dim TextRect As RECT 'This defines where it will be
    'Dim BorderColor As Long
    
    'Set width and height if no specified
    If width = 0 Then width = Len(Text) * (font_list(Font_Index).Size + 1)
    If height = 0 Then height = font_list(Font_Index).Size * 2
    
    'DrawBorder
    
    'BorderColor = D3DColorXRGB(0, 0, 0)
    
    'TextRect.top = top - 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left - 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top + 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left + 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + height
    TextRect.Right = left + width
    d3dx.DrawText font_list(Font_Index).dFont, Color, Text, TextRect, Alingment

End Sub
Private Function Font_Check(ByVal Font_Index As Long) As Boolean
    If Font_Index > 0 And Font_Index <= font_count Then
        Font_Check = True
    End If
End Function

Private Sub Fonts_Destroy()
    Dim i As Integer
    
    For i = 1 To font_count
        Set font_list(i).dFont = Nothing
        font_list(i).Size = 0
    Next i
    font_count = 0
End Sub

Public Function DXEngine_ParticleGroupCreate(ByVal map_x As Integer, ByVal map_y As Integer, ByVal particle_type As e_ParticleType, ByVal LifeTime As Integer, Optional ByVal char_index As Integer) As Integer
On Error GoTo ErrHandler

    If particle_type = 0 Then Exit Function
    DXEngine_ParticleGroupCreate = ParticleGroupMake(map_x, map_y, LifeTime, particle_type, char_index)
    
Exit Function
ErrHandler:
    DXEngine_ParticleGroupCreate = 0
End Function

Private Function ParticleGroupMake(ByVal map_x As Byte, ByVal map_y As Byte, ByVal LifeTime As Integer, ByVal particle_type As e_ParticleType, Optional ByVal char_index As Integer)

    
    Dim i As Integer
    Dim Particle_Index As Integer
    
    ParticleGroupMake = 0
    
    Call Particle_FreeIndexGet(Particle_Index)
    
    If Particle_Index = 0 Then
        particle_group_count = particle_group_count + 1
        Particle_Index = particle_group_count
        ReDim Preserve particle_group_list(1 To particle_group_count)
    End If
    
    particle_group_list(Particle_Index) = particle_group_data(particle_type)
    
    With particle_group_list(Particle_Index)
        .Active = 1
        .ParticleGroup_Lifetime = LifeTime
        .Created = GetTickCount
        .map_x = map_x
        .map_y = map_y
        .char_index = char_index
        
        If .ParticleEmisor.Delay1 Or .ParticleEmisor.Delay2 Then
            For i = 1 To .ParticleEmisor.Particle_count
                .ParticleEmisor.Particle(i).Delay = Val(General_Random_Number(.ParticleEmisor.Delay1, .ParticleEmisor.Delay2))
            Next i
        End If
    End With
    
    
    
    ParticleGroupMake = Particle_Index
End Function
Private Function Particles_Initialize()
    'particle_group_count = -1
End Function
Private Function Particles_GroupUpdate(ByVal ParticleGroup_Index As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal offset_x As Single, Optional ByVal offset_y As Single)
    Dim i As Integer
    Dim Time As Long
    Dim Particle_Emisor As tParticle_Emisor
    
    Time = GetTickCount
    With particle_group_list(ParticleGroup_Index)
        If ParticleGroup_CheckPermanency(ParticleGroup_Index, Time) Then
            With .ParticleEmisor
                If .Particle_Frame_Counter > .Particle_Speed Then
                    .Particle_Frame_Counter = 0
                    For i = 0 To .Particle_count
                        If .Particle(i).Created >= .Particle(i).Particle_LifeTime Then
                            .Particle(i).Alive = 0
                        ElseIf .KillWhenAtTarget Then
                            If .Particle(i).x + .Particle(i).Moved_X >= .Target_X Or .Particle(i).y + .Particle(i).Moved_Y >= .Target_Y Then
                                .Particle(i).Alive = 0
                            End If
                        End If
                    
                        If .Particle(i).Alive = 0 And Not .Particle(i).Used Then
                            If .Particle(i).DelayCounter >= .Particle(i).Delay Then
                                If .StartParticlesDestroy And Not .Particle(i).Used Then
                                    .Particle(i).Used = True
                                    .ParticlesLeft = .ParticlesLeft - 1
                                Else
                                    Select Case .ParticleGroup_Type
                                        Case 0
                                            Effect1_Reset ParticleGroup_Index, i
                                        Case 1
                                            Effect2_Reset ParticleGroup_Index, i
                                        Case 2
                                            Effect3_Reset ParticleGroup_Index, i
                                        Case 3
                                            Effect4_Reset ParticleGroup_Index, i
                                        Case 4
                                            Effect5_Reset ParticleGroup_Index, i
                                    End Select
                                    
                                    'Standard particle settings.
                                    .Particle(i).Alive = 1
                                    .Particle(i).Created = 0
                                    .Particle(i).vX = General_Random_Number(.vX1, .vX2)
                                    .Particle(i).vY = General_Random_Number(.vY1, .vY2)
                                    .Particle(i).Particle_LifeTime = General_Random_Number(.PLt1, .PLt2)
                                    
                                    'Reset moving status.
                                    .Particle(i).Moved_X = 0
                                    .Particle(i).Moved_Y = 0
    
                                    .Particle(i).texture_index = General_Random_Number(1, .texture_count)
                                
                                    .Particle(i).CurrentColor(1) = .StartColor(1)
                                    .Particle(i).CurrentColor(2) = .StartColor(2)
                                    .Particle(i).CurrentColor(3) = .StartColor(3)
                                    .Particle(i).CurrentColor(4) = .StartColor(4)
                                    
                                    .Particle(i).rgb_list(0) = D3DColorARGB(.Particle(i).CurrentColor(1).A, .Particle(i).CurrentColor(1).R, .Particle(i).CurrentColor(1).G, .Particle(i).CurrentColor(1).B)
                                    .Particle(i).rgb_list(1) = D3DColorARGB(.Particle(i).CurrentColor(2).A, .Particle(i).CurrentColor(2).R, .Particle(i).CurrentColor(2).G, .Particle(i).CurrentColor(2).B)
                                    .Particle(i).rgb_list(2) = D3DColorARGB(.Particle(i).CurrentColor(3).A, .Particle(i).CurrentColor(3).R, .Particle(i).CurrentColor(3).G, .Particle(i).CurrentColor(3).B)
                                    .Particle(i).rgb_list(3) = D3DColorARGB(.Particle(i).CurrentColor(4).A, .Particle(i).CurrentColor(4).R, .Particle(i).CurrentColor(4).G, .Particle(i).CurrentColor(4).B)
                                    
                                    
                                    .Particle(i).Screen_X = x + .Particle(i).x
                                    .Particle(i).Screen_Y = y + .Particle(i).y
                                End If
                            Else
                                .Particle(i).DelayCounter = .Particle(i).DelayCounter + 1
                            End If
                        Else
                            If Not .Particle(i).Used Then
                                Call Particle_Update(ParticleGroup_Index, i, x, y, Time, offset_x, offset_y)
                            End If
                        End If
                    Next i
                Else
                    .Particle_Frame_Counter = .Particle_Frame_Counter + particle_timer
                End If
            End With
        Else
            DXEngine_ParticleGroup_Destroy (ParticleGroup_Index)
        End If
    End With
End Function
Public Sub DXEngine_ParticleGroupRender(ByVal Particle_Group_Index As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal offset_x As Single, Optional ByVal offset_y As Single)
    Dim i As Integer
    Dim rgb_list(3) As Long
    Dim Size As Integer
    
    Call Particles_GroupUpdate(Particle_Group_Index, x, y, offset_x, offset_y)
    
    With particle_group_list(Particle_Group_Index).ParticleEmisor
        For i = 1 To .Particle_count
            'If is destroyed...
            If .Particle(i).Alive Then
                If Not .Particle(i).texture_index = 0 Then
                    Size = .texture_size(.Particle(i).texture_index)
                    DXEngine_TextureRender .texture_index(.Particle(i).texture_index), .Particle(i).Screen_X, .Particle(i).Screen_Y, Size, Size, .Particle(i).rgb_list(), 0, 0, Size, Size, .alpha_blend, .Particle(i).angle
                End If
            End If
        Next i
    End With
End Sub

Private Sub Particle_FreeIndexGet(ByRef Particle_Index As Integer)
    Dim i As Byte
    
    If particle_group_count = 0 Then
        Particle_Index = 0
        Exit Sub
    End If
    i = 1
    Do While particle_group_list(i).Active
        If i >= particle_group_count Then
            Particle_Index = 0
            Exit Sub
        Else
            i = i + 1
        End If
    Loop
    
    Particle_Index = i
End Sub
Public Sub DXEngine_ParticleGroup_Destroy(ByVal ParticleGroup_Index As Integer)
    With particle_group_list(ParticleGroup_Index)
        .Active = 0
        Engine.Char_ParticleGroup_Destroy .char_index
        Engine.Map_ParticleGroup_Destroy .map_x, .map_y
    End With
End Sub
Public Sub DXEngine_ParticleGroupDestroyAll()
    Dim i As Byte
    If particle_group_count = 0 Then Exit Sub 'Particles are already destroyed
    For i = 1 To particle_group_count
        With particle_group_list(i)
            DXEngine_ParticleGroup_Destroy (i)
        End With
    Next i
    particle_group_count = 0
End Sub

Public Function D3DColorValueGet(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As D3DCOLORVALUE
    D3DColorValueGet.A = A
    D3DColorValueGet.R = R
    D3DColorValueGet.G = G
    D3DColorValueGet.B = B
End Function
Public Function ParticleSpeedCalculate(ByVal timer_elapsed_time As Single)
    particle_timer = timer_elapsed_time * 0.03
End Function

Public Sub DXEngine_TextureToHdcRender(ByVal texture_index As Long, desthdc As Long, ByVal Screen_X As Long, ByVal Screen_Y As Long, ByVal SX As Integer, ByVal SY As Integer, ByVal sW As Integer, ByVal sH As Integer, Optional transparent As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/02/2003
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************

    Dim file_path As String
    Dim Src_X As Long
    Dim Src_Y As Long
    Dim src_width As Long
    Dim src_height As Long
    Dim hdcsrc As Long

    file_path = resource_path & "\graphics\" & texture_index & ".bmp"
    
    Src_X = SX
    Src_Y = SY
    src_width = sW
    src_height = sH

    hdcsrc = CreateCompatibleDC(desthdc)
    
    SelectObject hdcsrc, LoadPicture(file_path)
    
    If transparent = False Then
        BitBlt desthdc, Screen_X, Screen_Y, src_width, src_height, hdcsrc, Src_X, Src_Y, SRCCOPY
    Else
        TransparentBlt desthdc, Screen_X, Screen_Y, src_width, src_height, hdcsrc, Src_X, Src_Y, src_width, src_height, COLOR_KEY
    End If
        
    DeleteDC hdcsrc
End Sub

Public Sub DXEngine_BeginSecondaryRender()
    Device_Clear
    ddevice.BeginScene
End Sub
Public Sub DXEngine_EndSecondaryRender(ByVal hwnd As Long, ByVal width As Integer, ByVal height As Integer)
    Dim DR As RECT
    DR.left = 0
    DR.top = 0
    DR.bottom = height
    DR.Right = width
    
    ddevice.EndScene
    ddevice.Present DR, ByVal 0&, hwnd, ByVal 0&
End Sub

Private Sub LoadParticles()
    Dim particle_path As String
    Dim i As Integer
    Dim aux As String
    Dim j As Byte
    
    particle_path = resource_path & PATH_INIT & "\particles.ini"
    
    If Not General_File_Exists(particle_path, vbArchive) Then Exit Sub
    
    particle_group_data_count = General_Var_Get(particle_path, "INIT", "Total")
    
    ReDim particle_group_data(1 To particle_group_data_count)
    
    'On Error Resume Next
    
    For i = 1 To particle_group_data_count
        With particle_group_data(i)
            With .ParticleEmisor
                .ParticleGroup_Type = Val(General_Var_Get(particle_path, i, "Tipo"))
                
                .alpha_blend = CBool(General_Var_Get(particle_path, i, "AlphaBlend"))
                .Bounce_Strength = Val(General_Var_Get(particle_path, i, "Bounce_Strength"))
                .Bounce_Y = Val(General_Var_Get(particle_path, i, "BounceY"))
                If .Bounce_Y = 0 Then .Bounce_Y = 16
                
                .ColorVariation = Val(General_Var_Get(particle_path, i, "ColorVariation"))
                
                If .ColorVariation Then
                    For j = 1 To 4
                        aux = General_Var_Get(particle_path, i, "ColorSet" & j)
                        .StartColor(j) = D3DColorValueGet(Val(general_field_read(1, aux, Asc(","))), Val(general_field_read(2, aux, Asc(","))), general_field_read(3, aux, Asc(",")), Val(general_field_read(4, aux, Asc(","))))
                        aux = General_Var_Get(particle_path, i, "ColorEnd" & j)
                        .EndColor(j) = D3DColorValueGet(Val(general_field_read(1, aux, Asc(","))), Val(general_field_read(2, aux, Asc(","))), Val(general_field_read(3, aux, Asc(","))), Val(general_field_read(4, aux, Asc(","))))
                    Next j
                Else
                    For j = 1 To 4
                        aux = General_Var_Get(particle_path, i, "ColorSet" & j)
                        .StartColor(j) = D3DColorValueGet(Val(general_field_read(1, aux, Asc(","))), Val(general_field_read(2, aux, Asc(","))), Val(general_field_read(3, aux, Asc(","))), Val(general_field_read(4, aux, Asc(","))))
                    Next j
                End If
                
                .Particle_count = Val(General_Var_Get(particle_path, i, "NumOfParticles"))
                .Friction = Val(General_Var_Get(particle_path, i, "Friction"))
                .PLt1 = Val(General_Var_Get(particle_path, i, "Life1"))
                .PLt2 = Val(General_Var_Get(particle_path, i, "Life2"))
                
                'Accelerations
                .Wind = Val((General_Var_Get(particle_path, i, "Wind")))
                .Gravity = Val(General_Var_Get(particle_path, i, "Gravity"))
                
                'Rotation
                .Spin = CByte(General_Var_Get(particle_path, i, "Spin"))
                .SpinH = Val(General_Var_Get(particle_path, i, "Spin_SpeedH"))
                .SpinL = Val(General_Var_Get(particle_path, i, "Spin_SpeedL"))
                
                .texture_count = Val(General_Var_Get(particle_path, i, "NumGrhs"))
                If .texture_count > 0 Then
                    ReDim .texture_index(1 To .texture_count) As Integer
                    ReDim .texture_size(1 To .texture_count) As Integer
                    For j = 1 To .texture_count
                        .texture_index(j) = Val(general_field_read(j, General_Var_Get(particle_path, i, "Grh_List"), Asc(",")))
                        .texture_size(j) = Val(general_field_read(j, General_Var_Get(particle_path, i, "Size_List"), Asc(",")))
                    Next j
                End If
                
                .vX1 = Val(General_Var_Get(particle_path, i, "VecX1"))
                .vX2 = Val(General_Var_Get(particle_path, i, "VecX2"))
                .vY1 = Val(General_Var_Get(particle_path, i, "VecY1"))
                .vY2 = Val(General_Var_Get(particle_path, i, "VecY2"))
                
                'Particle Startup position
                .x1 = Val(General_Var_Get(particle_path, i, "X1"))
                .x2 = Val(General_Var_Get(particle_path, i, "X2"))
                .Y1 = Val(General_Var_Get(particle_path, i, "Y1"))
                .Y2 = Val(General_Var_Get(particle_path, i, "Y2"))
                
                'Speed
                .Particle_Speed = Val(General_Var_Get(particle_path, i, "Speed"))
                If .Particle_Speed = 0 Then .Particle_Speed = 0.5
                
                'For circle Effects
                .Ratio = Val(General_Var_Get(particle_path, i, "Radio"))
                .RatioFriction = Val(General_Var_Get(particle_path, i, "RatioFriction"))
                .RatioVariation = Val(General_Var_Get(particle_path, i, "RatioVariation"))
                
                'Shaking
                .MoveX = Val(General_Var_Get(particle_path, i, "XMove"))
                .MoveY = Val(General_Var_Get(particle_path, i, "YMove"))
                
                .move_x1 = Val(General_Var_Get(particle_path, i, "move_x1"))
                .move_y1 = Val(General_Var_Get(particle_path, i, "move_y1"))
                .move_x2 = Val(General_Var_Get(particle_path, i, "move_x2"))
                .move_y2 = Val(General_Var_Get(particle_path, i, "move_y2"))
                
                'Variables de efectos especiales
                .TargetRatio = Val(General_Var_Get(particle_path, i, "TargetRatio"))
                .StopAtTargetRatio = Val((General_Var_Get(particle_path, i, "StopAtTargetRatio")))
                .KillWhenAtTarget = Val((General_Var_Get(particle_path, i, "KillWhenAtTarget")))
                .Target_X = Val(General_Var_Get(particle_path, i, "TargetX"))
                .Target_Y = Val(General_Var_Get(particle_path, i, "TargetY"))
                
                'Delay
                .Delay1 = Val(General_Var_Get(particle_path, i, "Delay1"))
                .Delay2 = Val(General_Var_Get(particle_path, i, "Delay2"))
                
                .CurrentRatio = .Ratio
                .ParticlesLeft = .Particle_count
                ReDim .Particle(0 To .Particle_count)
            End With
        End With
    Next i
End Sub

Public Sub DXEngine_DrawBox(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal Color As Long, Optional ByVal border_width = 1)
    Dim DR As RECT
    Dim VertexB(3) As TLVERTEX
    Dim box_rect As RECT
    
    With box_rect
        .bottom = y + height
        .left = x
        .Right = x + width
        .top = y
    End With
    
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        
    ddevice.SetTexture 0, Nothing
    
    'Upper Line
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top + border_width, 0, 1, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top + border_width, 0, 1, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Left Line
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.bottom, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 2, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 2, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Right Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.top, 0, 3, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.bottom, 0, 3, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Bottom Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom - border_width, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom - border_width, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 1, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub
Public Sub D3DColorToRgbList(rgb_list() As Long, Color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(Color.A, Color.R, Color.G, Color.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub
Private Sub Effect1_Reset(GroupIndex As Integer, ParticleIndex As Integer)
    With particle_group_list(GroupIndex).ParticleEmisor
        .Particle(ParticleIndex).x = General_Random_Number(.x1, .x2)
        .Particle(ParticleIndex).y = General_Random_Number(.Y1, .Y2)
    End With
End Sub
Private Sub Effect2_Reset(GroupIndex As Integer, ParticleIndex As Integer)
    Dim angle As Single

    With particle_group_list(GroupIndex).ParticleEmisor
        angle = ParticleIndex * (360 / .Particle_count) * DegreeToRadian

        .Particle(ParticleIndex).x = .x1 - (Sin(angle) * .CurrentRatio)
        .Particle(ParticleIndex).y = .Y1 + (Cos(angle) * .CurrentRatio)
    End With
End Sub
Private Sub Effect3_Reset(GroupIndex As Integer, ParticleIndex As Integer)
    Dim R As Single
    Dim angle As Single
    
    With particle_group_list(GroupIndex).ParticleEmisor
        angle = 360 / .Particle_count * ParticleIndex * DegreeToRadian
        R = Rnd

        .Particle(ParticleIndex).x = .x1 - (Sin(angle) * (R * .CurrentRatio))
        .Particle(ParticleIndex).y = .Y1 + (Cos(angle) * (R * .CurrentRatio))
    End With
End Sub
Private Sub Effect4_Reset(GroupIndex As Integer, ParticleIndex As Integer)
    Dim R As Single

    With particle_group_list(GroupIndex).ParticleEmisor
        R = Sin(20 / (ParticleIndex + 1)) * .CurrentRatio
    
        .Particle(ParticleIndex).x = R * Cos(ParticleIndex)
        .Particle(ParticleIndex).y = R * Sin(ParticleIndex)
    End With
End Sub

Private Sub Effect5_Reset(GroupIndex As Integer, ParticleIndex As Integer)
    Dim R As Single

    With particle_group_list(GroupIndex).ParticleEmisor
        R = .CurrentRatio + Rnd * 15 * Cos(2 * ParticleIndex)
        
        .Particle(ParticleIndex).x = R * Cos(ParticleIndex)
        .Particle(ParticleIndex).y = R * Sin(ParticleIndex)
    End With
End Sub

Private Sub Particle_Update(ByVal GroupIndex As Integer, ByVal i As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Time As Long, Optional ByVal offset_x As Single, Optional ByVal offset_y As Single)
    Dim angle As Single 'Used to calculate ratio variation.
    
    'Change color
    With particle_group_list(GroupIndex)
        With .ParticleEmisor
            .Particle(i).Created = .Particle(i).Created + 1
            
            If .ColorVariation Then
                Call D3DXColorLerp(.Particle(i).CurrentColor(1), .StartColor(1), .EndColor(1), (.Particle(i).Created) / .Particle(i).Particle_LifeTime)
                Call D3DXColorLerp(.Particle(i).CurrentColor(2), .StartColor(2), .EndColor(2), (.Particle(i).Created) / .Particle(i).Particle_LifeTime)
                Call D3DXColorLerp(.Particle(i).CurrentColor(3), .StartColor(3), .EndColor(3), (.Particle(i).Created) / .Particle(i).Particle_LifeTime)
                Call D3DXColorLerp(.Particle(i).CurrentColor(4), .StartColor(4), .EndColor(4), (.Particle(i).Created) / .Particle(i).Particle_LifeTime)
                
                .Particle(i).rgb_list(0) = D3DColorARGB(.Particle(i).CurrentColor(1).A, .Particle(i).CurrentColor(1).R, .Particle(i).CurrentColor(1).G, .Particle(i).CurrentColor(1).B)
                .Particle(i).rgb_list(1) = D3DColorARGB(.Particle(i).CurrentColor(2).A, .Particle(i).CurrentColor(2).R, .Particle(i).CurrentColor(2).G, .Particle(i).CurrentColor(2).B)
                .Particle(i).rgb_list(2) = D3DColorARGB(.Particle(i).CurrentColor(3).A, .Particle(i).CurrentColor(3).R, .Particle(i).CurrentColor(3).G, .Particle(i).CurrentColor(3).B)
                .Particle(i).rgb_list(3) = D3DColorARGB(.Particle(i).CurrentColor(4).A, .Particle(i).CurrentColor(4).R, .Particle(i).CurrentColor(4).G, .Particle(i).CurrentColor(4).B)
            End If

            'Do Shaking
            If .MoveX Then .Particle(i).vX = General_Random_Number(.move_x1, .move_x2)
            If .MoveY Then .Particle(i).vX = General_Random_Number(.move_y1, .move_y2)
                                                                                                      
            'Do Gravity
            If .Gravity Then
                .Particle(i).vY = .Particle(i).vY + .Gravity
            End If
                                
            If .Bounce_Strength <> 0 Then
                If .Particle(i).y + .Particle(i).Moved_Y > .Bounce_Y Then
                    .Particle(i).vY = .Bounce_Strength
                End If
            End If
            
            If .Spin Then .Particle(i).angle = .Particle(i).angle + General_Random_Number(.SpinL, .SpinH) / 100
                                   
            If .Wind Then
                .Particle(i).vX = .Particle(i).vX + (.Wind / .RatioFriction)
            End If
                        
            If .RatioVariation <> 0 Then
                .CurrentRatio = .CurrentRatio + .RatioVariation / .Friction
                angle = i * (360 / .Particle_count) * DegreeToRadian
                .Particle(i).x = .x1 - (Sin(angle) * .CurrentRatio)
                .Particle(i).y = .Y1 + (Cos(angle) * .CurrentRatio)
                
                If .StopAtTargetRatio Then
                     If .CurrentRatio >= .TargetRatio Then
                        .RatioVariation = 0 'Stop variation
                        .CurrentRatio = .TargetRatio
                     End If
                Else
                    If .CurrentRatio >= .TargetRatio Then
                        .CurrentRatio = .Ratio
                    End If
                End If
            End If
            
            'Move our particle
            .Particle(i).Moved_X = .Particle(i).Moved_X + .Particle(i).vX / .Friction + offset_x
            .Particle(i).Moved_Y = .Particle(i).Moved_Y + .Particle(i).vY / .Friction + offset_y
            
            .Particle(i).Screen_X = x + .Particle(i).x + .Particle(i).Moved_X
            .Particle(i).Screen_Y = y + .Particle(i).y + .Particle(i).Moved_Y
        End With
    End With
End Sub

Private Function ParticleGroup_CheckPermanency(ByVal ParticleGroup_Index As Integer, ByVal Time As Long) As Boolean
    
    ParticleGroup_CheckPermanency = False
    
    With particle_group_list(ParticleGroup_Index)
        If .ParticleGroup_Lifetime > Time - .Created Or .ParticleGroup_Lifetime = -1 And Not .ParticleEmisor.StartParticlesDestroy Then
            If .ParticleEmisor.ParticlesLeft > 0 Then
                ParticleGroup_CheckPermanency = True
            End If
        Else
            If .ParticleEmisor.ParticlesLeft > 0 Then
                ParticleGroup_CheckPermanency = True
                .ParticleEmisor.StartParticlesDestroy = True
            End If
        End If
    End With
End Function

Public Function DXEngine_ParticleGroup_End(ByVal ParticleGroup_Index As Integer)
    With particle_group_list(ParticleGroup_Index)
        .ParticleEmisor.StartParticlesDestroy = True
    End With
End Function

