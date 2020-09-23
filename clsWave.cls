VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Initialize()
If IsInitialized = True Then
Set DS = DX.DirectSoundCreate("")
DS.SetCooperativeLevel ThehWnd, DSSCL_PRIORITY
If Err.Number <> DD_OK Then MsgBox "Error initializing DirectSound", vbExclamation, "Error"
End If
End Sub

Public Sub StartWave(Filename As String, FileNumber As Integer)
If IsInitialized = True Then
Dim BufferDesc As DSBUFFERDESC
Dim DSStepFormat As WAVEFORMATEX
BufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
DSStepFormat.nFormatTag = WAVE_FORMAT_PCM
DSStepFormat.nChannels = 2
DSStepFormat.lSamplesPerSec = 22050
DSStepFormat.nBitsPerSample = 16
DSStepFormat.nBlockAlign = DSStepFormat.nBitsPerSample / 8 * DSStepFormat.nChannels
DSStepFormat.lAvgBytesPerSec = DSStepFormat.lSamplesPerSec * DSStepFormat.nBlockAlign
Set DSSound(i) = DS.CreateSoundBufferFromFile(Filename, BufferDesc, DSStepFormat)
End If
End Sub

Public Sub StopWave(FileNumber As Integer)
If IsInitialized = True Then
DSSound(i).Stop
DSSound(i).SetCurrentPosition 0
End If
End Sub
