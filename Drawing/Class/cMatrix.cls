VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMatrix"
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

Friend Sub Create(): GdipCreateMatrix mHandle: End Sub
Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeleteMatrix mHandle
    mHandle = 0
End Sub

Public Sub SetElements(m11 As Single, m12 As Single, m21 As Single, m22 As Single, dx As Single, dy As Single)
    GdipSetMatrixElements mHandle, m11, m12, m21, m22, dx, dy
End Sub

Public Function Clone() As cMatrix
    Dim Matrix As New cMatrix
    Dim H As Long
    Matrix.Release
    GdipCloneMatrix mHandle, H
    Matrix.Handle = H
    Set Clone = Matrix
End Function

Public Sub Multiply(Matrix As cMatrix, Order As MatrixOrder)
    GdipMultiplyMatrix mHandle, Matrix.Handle, Order
End Sub

Public Sub SetTranslate(ByVal offsetX As Single, ByVal offsetY As Single, Order As MatrixOrder)
    GdipTranslateMatrix mHandle, offsetX, offsetY, Order
End Sub

Public Sub SetScale(ByVal sx As Single, ByVal sy As Single, Order As MatrixOrder)
    GdipScaleMatrix mHandle, sx, sy, Order
End Sub

Public Sub SetRotate(ByVal Angle As Single, Order As MatrixOrder)
    GdipRotateMatrix mHandle, Angle, Order
End Sub

Public Sub SetShear(ByVal sx As Single, ByVal sy As Single, Order As MatrixOrder)
    GdipShearMatrix mHandle, sx, sy, Order
End Sub

Public Sub Invert()
    GdipInvertMatrix mHandle
End Sub

Public Sub TransformPoints(Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipTransformMatrixPoints mHandle, Ps(LBound(Ps)), Count
End Sub

Public Sub TransformVectors(Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipVectorTransformMatrixPoints mHandle, Ps(LBound(Ps)), Count
End Sub

Public Function GetElements() As Single()
    Dim Arr(0 To 5) As Single
    GdipGetMatrixElements mHandle, Arr(0)
    GetElements = Arr
End Function

Public Function IsInvertible() As Boolean
    Dim Result As Long
    GdipIsMatrixInvertible mHandle, Result
    IsInvertible = Result
End Function

Public Function IsEqual(Matrix As cMatrix) As Boolean
    Dim Result As Long
    GdipIsMatrixEqual mHandle, Matrix.Handle, Result
    IsEqual = Result
End Function
