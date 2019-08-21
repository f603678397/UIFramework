VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitmap"
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

Public Property Get PixelFormat() As GpPixelFormat: GdipGetImagePixelFormat mHandle, PixelFormat: End Property

Friend Sub Release(): GdipDisposeImage mHandle: End Sub

Public Sub FromFile(ByVal FileName As String)
    Release
    GdipCreateBitmapFromFile StrPtr(FileName), mHandle
End Sub

Public Sub FromGraphics(Graphics As cGraphics, ByVal Width As Long, ByVal Height As Long)
    Release
    GdipCreateBitmapFromGraphics Width, Height, Graphics.Handle, mHandle
End Sub

Public Sub FromScan0(ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, PixelFormat As GpPixelFormat, Scan0() As Byte)
    If (Not Scan0) = -1 Then Exit Sub
    Release
    GdipCreateBitmapFromScan0 Width, Height, stride, PixelFormat, Scan0(LBound(Scan0)), mHandle
End Sub

Public Sub CreateEmptyBitmap(ByVal Width As Long, ByVal Height As Long)
    Release
    GdipCreateBitmapFromScan0 Width, Height, 0, PixelFormat32bppARGB, ByVal 0, mHandle
End Sub

Public Function Clone() As cBitmap
    Dim Bitmap As New cBitmap
    GdipCloneImage mHandle, Bitmap.Handle
    Set Clone = Bitmap
End Function

Public Function CloneArea(ByVal X As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single) As cBitmap
    Dim Bitmap As New cBitmap
    GdipCloneBitmapArea X, Y, W, H, Me.PixelFormat, mHandle, Bitmap.Handle
    Set CloneArea = Bitmap
End Function

Public Function CreateCacheBitmap(Graphics As cGraphics) As cCachedBitmap
    Dim Cache As cCachedBitmap
    GdipCreateCachedBitmap mHandle, Graphics.Handle, Cache.Handle
    Set CreateCacheBitmap = Cache
End Function

Public Function GetPixel(ByVal X As Long, ByVal Y As Long) As Long
    GdipBitmapGetPixel mHandle, X, Y, GetPixel
End Function

Public Sub SetPixel(ByVal X As Long, ByVal Y As Long, ByVal nColor As Long)
    GdipBitmapSetPixel mHandle, X, Y, nColor
End Sub

Public Sub ConvertFormat(nFormat As GpPixelFormat, dType As DitherType, pType As PaletteType, Palette As ColorPalette, AlphaPercent As Single)
    GdipBitmapConvertFormat mHandle, nFormat, dType, pType, Palette, AlphaPercent
End Sub

Public Sub SetResolution(ByVal xDpi As Single, ByVal yDpi As Single)
    GdipBitmapSetResolution mHandle, xDpi, yDpi
End Sub

Public Function GetHistogramSize(nFormat As HistogramFormat) As Long
    GdipBitmapGetHistogramSize nFormat, GetHistogramSize
End Function

Public Sub GetHistogramChannel0(nFormat As HistogramFormat, outChannel0() As Long, outChannel1() As Long, outChannel2() As Long, outChannel3() As Long)
    Dim nSize As Long
    Dim c0() As Long, c1() As Long, c2() As Long, c3() As Long
    nSize = Me.GetHistogramSize(HistogramFormatARGB)
    ReDim c0(nSize - 1) As Long
    ReDim c1(nSize - 1) As Long
    ReDim c2(nSize - 1) As Long
    ReDim c3(nSize - 1) As Long
    GdipBitmapGetHistogram mHandle, nFormat, nSize, c0(0), c1(0), c2(0), c3(0)
    outChannel0 = c0
    outChannel1 = c1
    outChannel2 = c2
    outChannel3 = c3
End Sub

Public Function GetBitmapData() As Long()
    Dim mRect As RECTL
    Dim data() As Long
    Dim BmpData As BitmapData
    Dim Format As GpPixelFormat
    
    mRect.Right = Me.Width
    mRect.Bottom = Me.Height
    Format = Me.PixelFormat
    
    ReDim data(mRect.Right * mRect.Bottom - 1) As Long
    
    BmpData.Width = mRect.Right
    BmpData.Height = mRect.Bottom
    BmpData.PixelFormat = Format
    BmpData.Scan0 = VarPtr(data(0))
    BmpData.stride = 4 * mRect.Right
    
    GdipBitmapLockBits mHandle, mRect, ImageLockModeUserInputBuf Or ImageLockModeRead, Format, BmpData
    GdipBitmapUnlockBits mHandle, BmpData
    GetBitmapData = data
End Function

Public Function SetBitmapData(nColor() As Long) As Boolean
    On Error GoTo Err
    Dim mRect As RECTL
    Dim BmpData As BitmapData
    Dim Format As GpPixelFormat
    
    mRect.Right = Me.Width
    mRect.Bottom = Me.Height
    Format = Me.PixelFormat
    
    BmpData.Width = mRect.Right
    BmpData.Height = mRect.Bottom
    BmpData.PixelFormat = Format
    BmpData.Scan0 = VarPtr(nColor(0))
    BmpData.stride = 4 * mRect.Right
    
    GdipBitmapLockBits mHandle, mRect, ImageLockModeUserInputBuf Or ImageLockModeWrite, Format, BmpData
    GdipBitmapUnlockBits mHandle, BmpData
    SetBitmapData = True
Err:
End Function

Public Sub ApplyEffect(Effect As cEffect, mRect As RECTL)
    GdipBitmapApplyEffect mHandle, Effect.Handle, mRect, 0, 0, 0
End Sub

Public Function ApplyEffectNewBitmap(Effect As cEffect, mRect As RECTL, outRect As RECTL) As cBitmap
    Dim Bitmap As New cBitmap
    GdipBitmapCreateApplyEffect mHandle, 1, Effect.Handle, mRect, outRect, Bitmap.Handle, 0, 0, 0
    Set ApplyEffectNewBitmap = Bitmap
End Function

Public Function GetGraphics() As cGraphics
    Dim Graphics As New cGraphics
    GdipGetImageGraphicsContext mHandle, Graphics.Handle
    Set GetGraphics = Graphics
End Function

Public Function ToImage() As cImage
    Dim Img As New cImage
    Img.Handle = mHandle
    Set ToImage = Img
End Function