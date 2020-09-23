VERSION 5.00
Begin VB.UserControl Engine 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Engine.ctx":0000
   ScaleHeight     =   750
   ScaleWidth      =   900
   ToolboxBitmap   =   "Engine.ctx":236A
   Begin VB.PictureBox OpenPic 
      AutoSize        =   -1  'True
      Height          =   135
      Left            =   0
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   240
      Picture         =   "Engine.ctx":267C
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Surface As New clsSurface
Public Wave As New clsWave
Public MIDI As New clsMIDI

Public Sub About()
frmAbout.Show 1
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 750
UserControl.Width = 900
End Sub

'Public Function Initialize(FormHolder As Form, PicBox As PictureBox, Width As Integer, Height As Integer) As Long
Public Sub Initialize(FormID As Object, PicBox As Object, Width As Integer, Height As Integer)
Set DD = DX.DirectDrawCreate("")
ThehDC = PicBox.hDC
Set TheForm = FormID
Set PicB0X = PicBox
ThehWnd = PicBox.hWnd
GameWidth = Width
GameHeight = Height
SetRect GlobalRect, 0, 0, Width, Height
Set DD = DX.DirectDrawCreate("")
DD.SetCooperativeLevel UserControl.hWnd, DDSCL_NORMAL
DescPrimary.lFlags = DDSD_CAPS
DescPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
Set DDPrimary = DD.CreateSurface(DescPrimary)
DescGetReady.lFlags = DDSD_CAPS
DescGetReady.lWidth = Width
DescGetReady.lHeight = Height
DescGetReady.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
SavePicture Picture1.Picture, App.Path & "\Temp.bmp"
Set DDGetReady = DD.CreateSurfaceFromFile(App.Path & "\Temp.bmp", DescGetReady)
DDTransparent.high = vbBlack
DDTransparent.low = vbBlack
DDGetReady.SetColorKey DDCKEY_SRCBLT, DDTransparent
Kill App.Path & "\Temp.bmp"
Set DDClipper = DD.CreateClipper(0)
DDClipper.SetHWnd PicBox.hWnd
DDPrimary.SetClipper DDClipper
IsInitialized = True
End Sub

Public Sub DrawText(Text As String, X As Integer, y As Integer, Shadow As Boolean, ColourRed As Integer, ColourGreen As Integer, ColourBlue As Integer)
If IsInitialized = True Then
If Shadow = True Then
DDGetReady.SetForeColor vbBlack
DDGetReady.DrawText X + 1.5, y + 1.5, Text, False
DDGetReady.SetForeColor RGB(ColourRed, ColourGreen, ColourBlue)
DDGetReady.DrawText X, y, Text, False
Else
DDGetReady.SetForeColor RGB(ColourRed, ColourGreen, ColourBlue)
DDGetReady.DrawText X, y, Text, False
End If
End If
End Sub

Public Sub SetFont(Name As String, Size As Integer, Optional Bold As Boolean = False, Optional Italic As Boolean = False, Optional Underline As Boolean = False, Optional Strikethrough As Boolean = False)
If IsInitialized = True Then
TheFont.Name = Name
TheFont.Size = Size
If Bold = True Then TheFont.Bold = True
If Italic = True Then TheFont.Italic = True
If Underline = True Then TheFont.Underline = True
If Strikethrough = True Then TheFont.Strikethrough = True
DDGetReady.SetFont TheFont
End If
End Sub




