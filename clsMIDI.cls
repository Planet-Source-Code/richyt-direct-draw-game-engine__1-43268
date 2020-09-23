VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMIDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Initialize()
If IsInitialized = True Then
On Error Resume Next
Set DDMLoad = DX.DirectMusicLoaderCreate
Set DDMPerf = DX.DirectMusicPerformanceCreate
DDMPerf.Init Nothing, ThehWnd
DDMPerf.SetPort -1, 1
DDMPerf.SetMasterAutoDownload True
If Err.Number <> DD_OK Then MsgBox "Error initializing DirectMusic", vbExclamation, "Error"
End If
End Sub

Public Sub LoadMIDI(Filename As String)
If IsInitialized = True Then
On Error Resume Next
Set DDMSeg = DDMLoad.LoadSegment(Filename)
If Err.Number <> DD_OK Then MsgBox "ERROR : Could not load MIDI file!", vbExclamation, "ERROR!"
End If
End Sub

Public Sub PlayMIDI()
If IsInitialized = True Then
On Error Resume Next
DDMPerf.SetMasterVolume 5
DDMPerf.PlaySegment DDMSeg, 0, 0
End If
End Sub

Public Sub StopMIDI()
If IsInitialized = True Then
On Error Resume Next
DDMPerf.Stop DDMSeg, Nothing, 0, 0
End If
End Sub
