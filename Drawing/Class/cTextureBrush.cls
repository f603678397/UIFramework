VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTextureBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long): mHandle = nHandle: End Property

Public Property Get FillWrapMode() As WrapMode: GdipGetTextureWrapMode mHandle, FillWrapMode: End Property
Public Property Let FillWrapMode(Mode As WrapMode): GdipSetTextureWrapMode mHandle, Mode: End Property

Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeleteBrush mHandle
    mHandle = 0
End Sub

Public Sub FromBitmap(Bitmap As cBitmap, Flip As WrapMode)
    Release
    GdipCreateTexture Bitmap.Handle, Flip, mHandle
End Sub

Public Sub FromImage(Image As cImage, Flip As WrapMode)
    Release
    GdipCreateTexture Image.Handle, Flip, mHandle
End Sub

Public Sub FromBitmapArea(Bitmap As cBitmap, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Flip As WarpMode)
    Release
    GdipCreateTexture2 Bitmap.Handle, Flip, X, Y, Width, Height, mHandle
End Sub

Public Sub FromImageArea(Image As cImage, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Flip As WarpMode)
    Release
    GdipCreateTexture2 Image.Handle, Flip, X, Y, Width, Height, mHandle
End Sub

Public Sub FromBitmapWithAttr(Bitmap As cBitmap, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Attr As cImageAttributes)
    GdipCreateTextureIA Bitmap.Handle, Attr.Handle, X, Y, Width, Height, mHandle
End Sub

Public Sub FromImageWithAttr(Image As cImage, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Attr As cImageAttributes)
    Release
    GdipCreateTextureIA Image.Handle, Attr.Handle, X, Y, Width, Height, mHandle
End Sub

Public Function Clone() As cTextureBrush
    Dim Brush As New cTextureBrush
    Dim H As Long
    GdipCloneBrush mHandle, H
    Brush.Handle = H
    Set Clone = Brush
End Function

Public Function GetTransform() As cMatrix
    Dim Matrix As New cMatrix
    Dim H As Long
    Matrix.Release
    GdipGetTextureTransform mHandle, H
    Matrix.Handle = H
    Set GetTransform = Matrix
End Function

Public Sub SetTransform(Matrix As cMatrix)
    GdipSetTextureTransform mHandle, Matrix.Handle
End Sub

Public Sub ResetTransform()
    GdipResetTextureTransform mHandle
End Sub

Public Sub MultiplyTransform(Matrix As cMatrix, Optional Order As MatrixOrder = MatrixOrder.MatrixOrderPrepend)
    GdipMultiplyTextureTransform mHandle, Matrix.Handle, Order
End Sub

Public Sub SetTranslate(ByVal dx As Single, ByVal dy As Single, Optional Order As MatrixOrder = MatrixOrder.MatrixOrderPrepend)
    GdipTranslateTextureTransform mHandle, dx, dy, Order
End Sub

Public Sub SetScale(ByVal sx As Single, ByVal sy As Single, Optional Order As MatrixOrder = MatrixOrder.MatrixOrderPrepend)
    GdipScaleTextureTransform mHandle, sx, sy, Order
End Sub

Public Sub SetRotate(ByVal Angle As Single, Optional Order As MatrixOrder = MatrixOrder.MatrixOrderPrepend)
    GdipRotateTextureTransform mHandle, Angle, Order
End Sub

Public Function GetImage() As cImage
    Dim Img As New cImage
    Dim H As Long
    GdipGetTextureImage mHandle, H
    Img.Handle = H
    Set GetImage = Img
End Function

Public Function GetBitmap() As cBitmap
    Dim Bmp As New cBitmap
    Dim H As Long
    GdipGetTextureImage mHandle, H
    Bmp.Handle = H
    Set GetBitmap = Bmp
End Function

Public Function GetBaseBrush() As cBrush
    Dim Brush As New cBrush
    Dim H As Long
    GdipCloneBrush mHandle, H
    Brush.Handle = H
    Set GetBaseBrush = Brush
End Function
