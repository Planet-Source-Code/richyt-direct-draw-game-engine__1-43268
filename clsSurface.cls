VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Surface_Number As Integer

Public Function Add(Graphic As StdPicture, Name As String, StartX As Integer, StartY As Integer, Width As Integer, Height As Integer, Columns As Integer, Rows As Integer, Transparent_Red As Integer, Transparent_Green As Integer, Transparent_Blue As Integer, Optional This_Returns_The_Number_Of_The_Surface) As Integer
If IsInitialized = True Then
With SurfaceA(SurfaceCount)
.Blue = Transparent_Blue
.Column = 1
.Columns = Columns
Set .Graphic = Graphic
.Green = Transparent_Green
TheForm.Picture = .Graphic
Set TheForm.Picture = Nothing
.Height = Height
.Name = Name
.Red = Transparent_Red
.Row = 1
.Rows = Rows
.Width = Width
.X = StartX
.y = StartY
DescSurface.lFlags = DDSD_CAPS
DescSurface.lWidth = .Width
DescSurface.lHeight = .Height
DescSurface.lAlphaBitDepth = 0.1
SavePicture .Graphic, App.Path & "\Temp.bmp"
DDTransparent.high = RGB(.Red, .Green, .Blue)
DDTransparent.low = DDTransparent.high
Set DDSurface(SurfaceCount) = DD.CreateSurfaceFromFile(App.Path & "\Temp.bmp", DescSurface)
Kill App.Path & "\Temp.bmp"
DDSurface(SurfaceCount).SetColorKey DDCKEY_SRCBLT, DDTransparent
Add = SurfaceCount
SurfaceCount = SurfaceCount + 1
End With
End If
End Function

Public Function Clear()
If IsInitialized = True Then
DDGetReady.BltColorFill GlobalRect, vbWhite
End If
End Function

Public Sub Draw()
If IsInitialized = True Then
Dim BltRect As RECT
Dim SrcRect As RECT
DX.GetWindowRect ThehWnd, BltRect
SetRect SrcRect, 0, 0, GameWidth, GameHeight
DDPrimary.Blt BltRect, DDGetReady, SrcRect, DDBLT_WAIT
End If
End Sub

Public Sub DrawOn(Name As String, X As Integer, y As Integer, Column As Integer, Row As Integer)
If IsInitialized = True Then
Dim DestRect As RECT
Dim SrcRect As RECT
Dim i As Integer
Surface_Number = FindSurface(Name)
If Surface_Number = 1001 Then
Exit Sub
Else
With SurfaceA(Surface_Number)
.Column = Column
.Row = Row
.X = X
.y = y
SetRect DestRect, .X, .y, .Width / .Columns, .Height / .Rows
SrcRect.Left = .Width / .Columns * .Column - (.Width / .Columns)
SrcRect.Right = .Width / .Columns * .Column
SrcRect.Top = .Height / .Rows * .Row - (.Height / .Rows)
SrcRect.Bottom = .Height / .Rows * .Row

'SetRect SrcRect, 0, 0, 32, 32
'SetRect DestRect, 0, 0, 32, 32

DDGetReady.Blt DestRect, DDSurface(Surface_Number), SrcRect, DDBLT_KEYSRC
End With
End If
End If
End Sub

Private Function FindSurface(Name As String) As Integer
FindSurface = 1001
For i = 0 To 1000
If SurfaceA(i).Name = Name Then
FindSurface = i
i = 1000
End If
Next i
End Function
