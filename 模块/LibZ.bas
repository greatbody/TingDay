Attribute VB_Name = "内核"
Option Explicit

'--------------------枚举类型-------------------------
Public Enum TBMPDetectionMethod
  dmPeaks = 0
  dmAutoCorrelation = 1
End Enum

Public Enum TCallbackMessage
  MsgStopAsync = 1
  MsgPlayAsync = 2
  MsgEnterLoopAsync = 4
  MsgExitLoopAsync = 8
  MsgEnterFadeAsync = 16
  MsgExitFadeAsync = 32
  MsgStreamBufferDoneAsync = 64
  MsgStreamNeedMoreDataAsync = 128
  MsgNextSongAsync = 256
  MsgStop = 65536
  MsgPlay = 131072
  MsgEnterLoop = 262144
  MsgExitLoop = 524288
  MsgEnterFade = 1048576
  MsgExitFade = 209715
  MsgStreamBufferDone = 4194304
  MsgStreamNeedMoreData = 8388608
  MsgNextSong = 16777216
  MsgWaveBuffer = 33554432
End Enum

Public Enum TFFTGraphHorizontalScale
  gsLogarithmic = 0
  gsLinear = 1
End Enum

Public Enum TFFTGraphParamID
  gpFFTPoints = 1
  gpGraphType
  gpWindow
  gpHorizontalScale
  gpSubgrid
  gpTransparency
  gpFrequencyScaleVisible
  gpDecibelScaleVisible
  gpFrequencyGridVisible
  gpDecibelGridVisible
  gpBgBitmapVisible
  gpBgBitmapHandle
  gpColor1
  gpColor2
  gpColor3
  gpColor4
  gpColor5
  gpColor6
  gpColor7
  gpColor8
  gpColor9
  gpColor10
  gpColor11
  gpColor12
  gpColor13
  gpColor14
  gpColor15
  gpColor16
End Enum

Public Enum TFFTGraphSize
  FFTGraphMinWidth = 100
  FFTGraphMinHeight = 60
End Enum

Public Enum TFFTGraphType
  gtLinesLeftOnTop = 0
  gtLinesRightOnTop
  gtAreaLeftOnTop
  gtAreaRightOnTop
  gtBarsLeftOnTop
  gtBarsRightOnTop
  gtSpectrum
End Enum

Public Enum TFFTWindow
  fwRectangular = 1
  fwHamming
  fwHann
  fwCosine
  fwLanczos
  fwBartlett
  fwTriangular
  fwGauss
  fwBartlettHann
  fwBlackman
  fwNuttall
  fwBlackmanHarris
  fwBlackmanNuttall
  fwFlatTop
End Enum

Public Enum TID3Version
  id3Version1 = 1
  id3Version2 = 2
End Enum

Public Enum TSeekMethod
  smFromBeginning = 1
  smFromEnd = 2
  smFromCurrentForward = 4
  smFromCurrentBackward = 8
End Enum

Public Enum TSettingID
  sidWaveBufferSize = 1
  sidAccurateLength = 2
  sidAccurateSeek = 3
  sidSamplerate = 4
  sidChannelNumber = 5
  sidBitPerSample = 6
  sidBigEndian = 7
End Enum

Public Enum TStreamFormat
  sfUnknown = 0
  sfMp3 = 1
  sfOgg = 2
  sfWav = 3
  sfPCM = 4
  sfFLAC = 5
  sfFLACOgg = 6
  sfAC3 = 7
  sfAutodetect = 1000
End Enum

Public Enum TTimeFormat
  tfMillisecond = 1
  tfSecond = 2
  tfHMS = 4
  tfSamples = 8
End Enum

Public Enum TWaveOutFormat
  format_invalid = 0
  format_11khz_8bit_mono = 1
  format_11khz_8bit_stereo = 2
  format_11khz_16bit_mono = 4
  format_11khz_16bit_stereo = 8
  format_22khz_8bit_mono = 16
  format_22khz_8bit_stereo = 32
  format_22khz_16bit_mono = 64
  format_22khz_16bit_stereo = 128
  format_44khz_8bit_mono = 256
  format_44khz_8bit_stereo = 512
  format_44khz_16bit_mono = 1024
  format_44khz_16bit_stereo = 2048
End Enum

Public Enum TWaveOutFunctionality
  supportPitchControl = 1
  supportPlaybackRateControl = 2
  supportVolumeControl = 4
  supportSeparateLeftRightVolume = 8
  supportSync = 16
  supportSampleAccuratePosition = 32
  supportDirectSound = 6
End Enum

'-------------------------结构体--------------------------------------
Public Type TEchoEffect
  nLeftDelay As Long
  nLeftSrcVolume As Long
  nLeftEchoVolume As Long
  nRightDelay As Long
  nRightSrcVolume As Long
  nRightEchoVolume As Long
End Type

Public Type TID3Info
  title As Long
  Artist As Long
  Album As Long
  Year As Long
  Comment As Long
  TrackNum As Long
  Genre As Long
End Type

Public Type TStreamHMSTime
  hour As Long
  minute As Long
  second As Long
  millisecond As Long
End Type

Public Type TStreamTime
  sec As Long
  ms As Long
  samples As Long
  hms As TStreamHMSTime
End Type

Public Type TStreamInfo
  SamplingRate As Long
  ChannelNumber As Long
  VBR As Long
  Bitrate As Long
  length As TStreamTime
  description As Long
End Type

Public Type TStreamLoadInfo
  NumberOfBuffers As Long
  NumberOfBytes As Long
End Type

Public Type TStreamStatus
  fPlay As Long
  fPause As Long
  fEcho As Long
  fEqualizer As Long
  fVocalCut As Long
  fSideCut As Long
  fChannelMix As Long
  fSlideVolume As Long
  nLoop As Long
  fReverse As Long
  nSongIndex As Long
  nSongsInQueue As Long
End Type

Public Type TWaveOutInfo
  ManufacturerID As Long
  ProductID As Long
  DriverVersion As Long
  Formats As Long
  Channels As Long
  Support As Long
  ProductName As String
End Type

'---------------------------函数声明-------------------------------
'创建ZPlayer
Public Declare Function zplay_CreateZPlay Lib "libzplay.dll" () As Long
'销毁ZPlayer
Public Declare Function zplay_DestroyZPlay Lib "libzplay.dll" (ByVal objptr As Long) As Long
'获取错误信息
Public Declare Function zplay_GetError Lib "libzplay.dll" (ByVal objptr As Long) As Long
'获取错误信息2
Public Declare Function zplay_GetErrorW Lib "libzplay.dll" (ByVal objptr As Long) As Long
'打开文件
Public Declare Function zplay_OpenFile Lib "libzplay.dll" (ByVal objptr As Long, ByVal sFileName As String, ByVal nFormat As TStreamFormat) As Long
'打开流
Public Declare Function zplay_OpenStream Lib "libzplay.dll" (ByVal objptr As Long, ByVal fBuffered As Long, ByVal fManaged As Long, sMemStream As Any, ByVal nStreamSize As Long, ByVal nFormat As TStreamFormat) As Long
'关闭文件
Public Declare Function zplay_Close Lib "libzplay.dll" (ByVal objptr As Long) As Long
'播放
Public Declare Function zplay_Play Lib "libzplay.dll" (ByVal objptr As Long) As Long
'停止
Public Declare Function zplay_Stop Lib "libzplay.dll" (ByVal objptr As Long) As Long
'暂停
Public Declare Function zplay_Pause Lib "libzplay.dll" (ByVal objptr As Long) As Long
'继续播放
Public Declare Function zplay_Resume Lib "libzplay.dll" (ByVal objptr As Long) As Long
'获取播放状态
Public Declare Sub zplay_GetStatus Lib "libzplay.dll" (ByVal objptr As Long, pStatus As TStreamStatus)
'获取播放位置
Public Declare Sub zplay_GetPosition Lib "libzplay.dll" (ByVal objptr As Long, ByRef pTime As TStreamTime)
'循环播放
Public Declare Function zplay_PlayLoop Lib "libzplay.dll" (ByVal objptr As Long, ByVal fFormatStartTime As TTimeFormat, ByRef pStartTime As TStreamTime, ByVal fFormatEndTime As TTimeFormat, ByRef pEndTime As TStreamTime, ByVal nNumOfCycles As Long, ByVal fContinuePlaying As Long) As Long
'定位
Public Declare Function zplay_Seek Lib "libzplay.dll" (ByVal objptr As Long, ByVal fFormat As TTimeFormat, ByRef pTime As TStreamTime, ByVal nMoveMethod As TSeekMethod) As Long
'获取标签信息
Public Declare Function zplay_LoadID3 Lib "libzplay.dll" (ByVal objptr As Long, ByVal nId3Version As TID3Version, pId3Info As TID3Info) As Long
'获取流信息
Public Declare Sub zplay_GetStreamInfo Lib "libzplay.dll" (ByVal objptr As Long, pInfo As TStreamInfo)
'获取音量
Public Declare Sub zplay_GetPlayerVolume Lib "libzplay.dll" (ByVal objptr As Long, pnLeftVolume As Long, pnRightVolume As Long)
'设置音量
Public Declare Function zplay_SetPlayerVolume Lib "libzplay.dll" (ByVal objptr As Long, ByVal pnLeftVolume As Long, ByVal pnRightVolume As Long) As Long
'设置频谱参数
Public Declare Function zplay_SetFFTGraphParam Lib "libzplay.dll" (ByVal objptr As Long, ByVal nParamID As TFFTGraphParamID, ByVal nParamValue As Long) As Long
'绘制频谱
Public Declare Function zplay_DrawFFTGraphOnHDC Lib "libzplay.dll" (ByVal objptr As Long, ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
