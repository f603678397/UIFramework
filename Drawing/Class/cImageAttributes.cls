VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

Private Sub Class_Initialize(): Create: End Sub
Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long): mHandle = nHandle: End Property

Friend Sub Create()
    GdipCreateImageAttributes mHandle
End Sub

Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDisposeImageAttributes mHandle
    mHandle = 0
End Sub

Public Function Clone() As cImageAttributes
    Dim Attr As New cImageAttributes
    Dim H As Long
    Attr.Release
    GdipCloneImageAttributes mHandle, H
    Attr.Handle = H
    Set Clone = Attr
End Function

Public Sub ReSet(Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipResetImageAttributes mHandle, ColorAdjustTypeDefault
End Sub

Public Sub SetToIdentity(nType As ColorAdjustType)
    GdipSetImageAttributesToIdentity mHandle, nType
End Sub

Public Sub SetColorMatrix(CM As ColorMatrix, Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesColorMatrix mHandle, nType, TRUEVALUE, CM, ByVal 0&, ColorMatrixFlagsDefault
End Sub

Public Sub ClearColorMatrix(Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesColorMatrix mHandle, nType, FALSEVALUE, ByVal 0&, ByVal 0&, ColorMatrixFlagsDefault
End Sub

Public Sub SetThreshold(ByVal Threshold As Single, Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesThreshold mHandle, nType, TRUEVALUE, Threshold
End Sub

Public Sub ClearThreshold(Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesThreshold mHandle, nType, TRUEVALUE, 0
End Sub

Public Sub SetGamma(ByVal Gamma As Single, Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesGamma mHandle, nType, TRUEVALUE, Gamma
End Sub

Public Sub ClearGamma(Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesGamma mHandle, nType, FALSEVALUE, 0
End Sub

Public Sub SetNoOp(Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesNoOp mHandle, nType, TRUEVALUE
End Sub

Public Sub ClearNoOp(Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesNoOp mHandle, nType, FALSEVALUE
End Sub

Public Sub SetColorKeys(ByVal Low As Long, ByVal Hight As Long, Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesColorKeys mHandle, nType, TRUEVALUE, Low, Hight
End Sub

Public Sub ClearColorKeys(Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesColorKeys mHandle, nType, FALSEVALUE, 0, 0
End Sub

Public Sub SetRemapTable(ByVal mSize As Long, Map As ColorMap, Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesRemapTable mHandle, nType, TRUEVALUE, mSize, Map
End Sub

Public Sub ClearRemapTable(Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    GdipSetImageAttributesRemapTable mHandle, nType, FALSEVALUE, 0, 0
End Sub

Public Sub SetImageAttributesWrapMode(Wrap As WrapMode, nColor As Long, Optional bClamp As BoolType = FALSEVALUE)
    GdipSetImageAttributesWrapMode mHandle, Wrap, nColor, bClamp
End Sub

Public Function GetAdjustedPalette(ByVal nType As ColorAdjustType) As ColorPalette
    GdipGetImageAttributesAdjustedPalette mHandle, GetAdjustedPalette, nType
End Function

Public Sub SetOpacity(ByVal Opacity As Single, Optional ByVal nType As ColorAdjustType = ColorAdjustTypeDefault)
    Dim CM As Drawing2D.ColorMatrix
    With CM
        .M(0, 0) = 1
        .M(1, 1) = 1
        .M(2, 2) = 1
        .M(3, 3) = Opacity
        .M(4, 4) = 1
    End With
    GdipSetImageAttributesColorMatrix mHandle, nType, TRUEVALUE, CM, ByVal 0&, nType
End Sub
