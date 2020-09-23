Attribute VB_Name = "modEngine"
Public Type typSurface
    Graphic As StdPicture
    Name As String
    X As Integer
    y As Integer
    Columns As Integer
    Column As Integer
    Rows As Integer
    Row As Integer
    Red As Integer
    Green As Integer
    Width As Integer
    Height As Integer
    Blue As Integer
End Type

Public DX As New DirectX7
Public DS As DirectSound
Public DD As DirectDraw7
Public IsInitialized As Boolean
Public DDMLoad As DirectMusicLoader
Public DDMPerf As DirectMusicPerformance
Public DDMSeg As DirectMusicSegment
Public DDClipper As DirectDrawClipper
Public DDPrimary As DirectDrawSurface7
Public DSSound(100) As DirectSoundBuffer
Public DescPrimary As DDSURFACEDESC2
Public DDGetReady As DirectDrawSurface7
Public DescGetReady As DDSURFACEDESC2
Public DDCoverUp As DirectDrawSurface7
Public DescCoverUp As DDSURFACEDESC2
Public SurfaceA(0 To 1000) As typSurface
Public DescSurface As DDSURFACEDESC2
Public DDSurface(0 To 1000) As DirectDrawSurface7
Public SurfaceCount As Integer
Public DDTransparent As DDCOLORKEY
Public GlobalRect As RECT
Public ThehDC As Long
Public TheFont As New StdFont
Public GameWidth As Integer
Public GameHeight As Integer
Public ThehWnd As Long
Public PicB0X As PictureBox
Public TheForm As Form
Public Function SetRect(TheRect As RECT, Left As Integer, Top As Integer, Width As Integer, Height As Integer)
TheRect.Left = Left
TheRect.Top = Top
TheRect.Right = TheRect.Left + Width
TheRect.Bottom = TheRect.Top + Height
End Function

