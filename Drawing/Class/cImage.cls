VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long): mHandle = nHandle: End Property

Public Property Get Width() As Long: GdipGetImageWidth mHandle, Width: End Property
Public Property Get Height() As Long: GdipGetImageHeight mHandle, Height: End Property

Public Property Get Flag() As Long: GdipGetImageFlags mHandle, Flag: End Property

Public Property Get PixelFormat() As GpPixelFormat: GdipGetImagePixelFormat mHandle, PixelFormat: End Property

Friend Sub Release(): GdipDisposeImage mHandle: End Sub

Public Sub FromFile(ByVal FileName As String)
    GdipLoadImageFromFile FileName, mHandle
End Sub

Public Function Clone() As cImage
    Dim Img As New cImage
    Img.Handle = mHandle
    Set Clone = Img
End Function

Public Sub SaveToFile(ByVal FileName As String, nType As GdipImageType)
    Dim Encode As CLSID
    Select Case nType
        Case GdipImageType.Bmp
            CLSIDFromString StrPtr(gdiplus.ImageEncoderBMP), Encode
        Case GdipImageType.EMF
            CLSIDFromString StrPtr(gdiplus.ImageEncoderEMF), Encode
        Case GdipImageType.GIF
            CLSIDFromString StrPtr(gdiplus.ImageEncoderGIF), Encode
        Case GdipImageType.ICO
            CLSIDFromString StrPtr(gdiplus.ImageEncoderICO), Encode
        Case GdipImageType.JPG
            CLSIDFromString StrPtr(gdiplus.ImageEncoderJPG), Encode
        Case GdipImageType.PNG
            CLSIDFromString StrPtr(gdiplus.ImageEncoderPNG), Encode
        Case GdipImageType.TIF
            CLSIDFromString StrPtr(gdiplus.ImageEncoderTIF), Encode
        Case GdipImageType.WMF
            CLSIDFromString StrPtr(gdiplus.ImageEncoderWMF), Encode
    End Select
    
    GdipSaveImageToFile mHandle, FileName, Encode, ByVal 0&
End Sub

Public Sub SaveAdd(Params As EncoderParameters)
    GdipSaveAdd mHandle, Params
End Sub

Public Sub SaveAddImage(Img As cImage, Params As EncoderParameters)
    GdipSaveAddImage mHandle, Img.Handle, Params
End Sub

Public Function GetGraphics() As cGraphics
    Dim Graphics As New cGraphics
    GdipGetImageGraphicsContext mHandle, Graphics.Handle
    Set GetGraphics = Graphics
End Function

Public Function GetBounds(Optional Unit As GpUnit) As RECTF
    GdipGetImageBounds mHandle, GetBounds, Unit
End Function

Public Function GetDimension() As SIZEF
    Dim w As Single, H As Single
    GdipGetImageDimension mHandle, w, H
    GetDimension.cx = w
    GetDimension.cy = H
End Function

Public Function GetType() As Image_Type
    GdipGetImageType mHandle, GetType
End Function

Public Function GetRawFormat() As CLSID
    GdipGetImageRawFormat mHandle, GetRawFormat
End Function

Public Function GetThumbnail(ByVal Width As Long, ByVal Height As Long) As cImage
    Dim Img As New cImage
    GdipGetImageThumbnail mHandle, Width, Height, Img.Handle, 0, 0
    Set GetThumbnail = Img
End Function

Public Function GetHorizontalResolution() As Single
    GdipGetImageHorizontalResolution mHandle, GetHorizontalResolution
End Function

Public Function GetVerticalResolution() As Single
    GdipGetImageVerticalResolution mHandle, GetVerticalResolution
End Function

Public Function GetFrameDimensionsCount() As Long
    GdipImageGetFrameDimensionsCount mHandle, GetFrameDimensionsCount
End Function

Public Function GetFrameDimensionsList() As CLSID()
    Dim nID() As CLSID
    Dim Count As Long
    Count = Me.GetFrameDimensionsCount
    ReDim nID(Count - 1) As CLSID
    GdipImageGetFrameDimensionsList mHandle, nID(0), Count
    GetFrameDimensionsList = nID
End Function

Public Function GetFrameCount(nID As CLSID) As Long
    GdipImageGetFrameCount mHandle, nID, GetFrameCount
End Function

Public Sub SelectActiveFrame(nID As CLSID, ByVal Index As Long)
    GdipImageSelectActiveFrame mHandle, nID, Index
End Sub

Public Sub RotateFlip(RF As RotateFlipType)
    GdipImageRotateFlip mHandle, RF
End Sub

Public Function ToBitmap() As cBitmap
    Dim Bmp As New cBitmap
    Dim H As Long
    GdipCloneImage mHandle, H
    Bmp.Handle = H
    Set ToBitmap = Bmp
End Function
