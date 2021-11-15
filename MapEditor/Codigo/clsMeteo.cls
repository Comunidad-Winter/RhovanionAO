VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMeteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This Object provides info about the Ambient.
'Handles the rain, day and night.

'Time info
Private mHour As Byte
Private mMinutes As Byte
Private mSeconds As Byte

Private RainParticleIndex As Integer
Private Rainning As Byte
'Light Info
Private MedioDia As D3DCOLORVALUE
Private MediaNoche As D3DCOLORVALUE
Private LuzActual As D3DCOLORVALUE
Private mediotarde As D3DCOLORVALUE
Private dia As D3DCOLORVALUE

Public Sub Initialize()
    MedioDia.R = 255
    MedioDia.G = 255
    MedioDia.B = 255
    
    mediotarde.R = 192
    mediotarde.G = 175
    mediotarde.B = 139
    
    dia.R = 228
    dia.G = 192
    dia.B = 168
    
    MediaNoche.R = 100
    MediaNoche.G = 100
    MediaNoche.B = 120
End Sub

Public Sub Set_Time(ByVal Hour As Byte, ByVal Minutes As Byte)
    mHour = Hour
    mMinutes = Minutes
    mSeconds = 0
    
    AmbientCalculate
End Sub
Private Sub AmbientCalculate()
If mHour >= 7 Then
Call D3DXColorLerp(LuzActual, MediaNoche, dia, mHour / 7)
End If

If mHour >= 12 Then
Call D3DXColorLerp(LuzActual, dia, MedioDia, mHour / 12)
End If

If mHour >= 18 Then
Call D3DXColorLerp(LuzActual, MedioDia, mediotarde, mHour / 18)
End If

If mHour >= 21 Then
Call D3DXColorLerp(LuzActual, mediotarde, MediaNoche, mHour / 21)
End If
    
End Sub
Public Sub Get_AmbientLight(Color As D3DCOLORVALUE)
    Color = LuzActual
End Sub

Public Sub MeteoParticle_Render()
    If Rainning Then _
        DXEngine_ParticleGroupRender RainParticleIndex, 0, 0
End Sub

Public Sub Rain_Toggle()
    If Rainning Then
        Rain_Stop
    Else
        Rain_Start
    End If
End Sub
Private Sub Rain_Stop()
    Rainning = 0
    DXEngine_ParticleGroup_Destroy RainParticleIndex
    RainParticleIndex = 0
End Sub
Private Sub Rain_Start()
    If Rainning Then Exit Sub
    
    Rainning = 1
    
    RainParticleIndex = DXEngine_ParticleGroupCreate(0, 0, e_ParticleType.Rain, -1)
End Sub

Public Sub Get_AmbientColorRGB(rgb_list() As Long)
    rgb_list(0) = D3DColorARGB(LuzActual.A, LuzActual.R, LuzActual.G, LuzActual.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

