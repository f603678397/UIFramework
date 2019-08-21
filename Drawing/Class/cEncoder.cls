VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEncoder"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public Property Get ImageEncoderSuffix() As String
'    ImageEncoderSuffix = gdiplus.ImageEncoderSuffix
'End Property

Public Property Get ImageEncoderBMP() As String
    ImageEncoderBMP = gdiplus.ImageEncoderBMP
End Property

Public Property Get ImageEncoderJPG() As String
    ImageEncoderJPG = gdiplus.ImageEncoderJPG
End Property

Public Property Get ImageEncoderGIF() As String
    ImageEncoderGIF = gdiplus.ImageEncoderGIF
End Property

Public Property Get ImageEncoderEMF() As String
    ImageEncoderEMF = gdiplus.ImageEncoderEMF
End Property

Public Property Get ImageEncoderWMF() As String
    ImageEncoderWMF = gdiplus.ImageEncoderWMF
End Property

Public Property Get ImageEncoderTIF() As String
    ImageEncoderTIF = gdiplus.ImageEncoderTIF
End Property

Public Property Get ImageEncoderPNG() As String
    ImageEncoderPNG = gdiplus.ImageEncoderPNG
End Property

Public Property Get ImageEncoderICO() As String
    ImageEncoderICO = gdiplus.ImageEncoderICO
End Property

Public Property Get EncoderCompression() As String
    EncoderCompression = gdiplus.EncoderCompression
End Property

Public Property Get EncoderColorDepth() As String
    EncoderColorDepth = gdiplus.EncoderColorDepth
End Property

Public Property Get EncoderScanMethod() As String
    EncoderScanMethod = gdiplus.EncoderScanMethod
End Property

Public Property Get EncoderVersion() As String
    EncoderVersion = gdiplus.EncoderVersion
End Property

Public Property Get EncoderRenderMethod() As String
    EncoderRenderMethod = gdiplus.EncoderRenderMethod
End Property

Public Property Get EncoderQuality() As String
    EncoderQuality = gdiplus.EncoderQuality
End Property

Public Property Get EncoderTransformation() As String
    EncoderTransformation = gdiplus.EncoderTransformation
End Property

Public Property Get EncoderLuminanceTable() As String
    EncoderLuminanceTable = gdiplus.EncoderLuminanceTable
End Property

Public Property Get EncoderChrominanceTable() As String
    EncoderChrominanceTable = gdiplus.EncoderChrominanceTable
End Property

Public Property Get EncoderSaveFlag() As String
    EncoderSaveFlag = gdiplus.EncoderSaveFlag
End Property

Public Property Get EncoderColorSpace() As String
    EncoderColorSpace = gdiplus.EncoderColorSpace
End Property

Public Property Get EncoderImageItems() As String
    EncoderImageItems = gdiplus.EncoderImageItems
End Property

Public Property Get EncoderSaveAsCMYK() As String
    EncoderSaveAsCMYK = gdiplus.EncoderSaveAsCMYK
End Property
