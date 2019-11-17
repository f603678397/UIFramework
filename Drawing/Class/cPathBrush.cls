VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPathBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

'Private Sub Class_Initialize(): Create: End Sub
Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long): mHandle = nHandle: End Property

Public Property Get CenterColor() As Long
    GdipGetPathGradientCenterColor mHandle, CenterColor
End Property

Public Property Let CenterColor(ByVal nColor As Long)
    GdipSetPathGradientCenterColor mHandle, nColor
End Property

Public Property Get FocusScaleX() As Single
    GdipGetPathGradientFocusScales mHandle, FocusScaleX, 0
End Property

Public Property Get FocusScaleY() As Single
    GdipGetPathGradientFocusScales mHandle, 0, FocusScaleY
End Property

Public Property Get GammaCorrection() As Boolean
    Dim Result As Long
    GdipGetPathGradientGammaCorrection mHandle, Result
    GammaCorrection = Result
End Property

Public Property Let GammaCorrection(ByVal bGamma As Boolean)
    GdipSetPathGradientGammaCorrection mHandle, IIf(bGamma, 1, 0)
End Property

Public Property Get FillWrapMode() As WrapMode
    GdipGetPathGradientWrapMode mHandle, FillWrapMode
End Property

Public Property Let FillWrapMode(WrapMd As WrapMode)
    GdipSetPathGradientWrapMode mHandle, WrapMd
End Property

Public Property Get PointCount() As Long
    GdipGetPathGradientPointCount mHandle, PointCount
End Property

Public Property Get SurroundColorCount() As Long
    GdipGetPathGradientSurroundColorCount mHandle, SurroundColorCount
End Property

Public Property Get BlendCount() As Long
    GdipGetPathGradientBlendCount mHandle, BlendCount
End Property

Public Property Get PresetBlendCount() As Long
    GdipGetPathGradientPresetBlendCount mHandle, PresetBlendCount
End Property

Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeleteBrush mHandle
    mHandle = 0
End Sub

Public Sub FromPoints(Points() As POINTF)
    Dim Count As Long
    If (Not Points) = -1 Then Exit Sub
    
    Release
    
    Count = UBound(Points) - LBound(Points) + 1
    GdipCreatePathGradient Points(LBound(Points)), Count, WrapModeTileFlipXY, mHandle
End Sub

Public Sub FromPath(Path As cPath)
    Release
    GdipCreatePathGradientFromPath Path.Handle, mHandle
End Sub

Public Function Clone() As cPathBrush
    Dim Brush As New cPathBrush
    Dim H As Long
    GdipCloneBrush mHandle, H
    Brush.Handle = H
    Set Clone = Brush
End Function

Public Sub SetPath(Path As cPath)
    GdipSetPathGradientPath mHandle, Path.Handle
End Sub

Public Function GetPath() As cPath
    Dim Path As New cPath
    Dim H As Long
    Path.Release
    GdipGetPathGradientPath mHandle, H
    Path.Handle = H
    Set GetPath = Path
End Function

Public Function GetCenterPoint() As POINTF
    GdipGetPathGradientCenterPoint mHandle, GetCenterPoint
End Function

Public Sub SetCenterPoint(ByVal X As Single, ByVal Y As Single)
    Dim pt As POINTF
    pt.X = X: pt.Y = Y
    GdipSetPathGradientCenterPoint mHandle, pt
End Sub

Public Sub SetFoucsScales(ByVal xScale As Single, ByVal yScale As Single)
    GdipSetPathGradientFocusScales mHandle, xScale, yScale
End Sub

Public Function GetRectangle() As RECTF
    GdipGetPathGradientRect mHandle, GetRectangle
End Function

Public Function GetTransform() As cMatrix
    Dim Matrix As New cMatrix
    Dim H As Long
    Matrix.Release
    GdipGetPathGradientTransform mHandle, H
    Matrix.Handle = H
    Set GetTransform = Matrix
End Function

Public Sub SetTransform(Matrix As cMatrix)
    GdipSetPathGradientTransform mHandle, Matrix.Handle
End Sub

Public Sub ResetTransform()
    GdipResetPathGradientTransform mHandle
End Sub

Public Sub SetRotate(ByVal fAngle As Single, Order As MatrixOrder)
    GdipRotatePathGradientTransform mHandle, fAngle, Order
End Sub

Public Sub SetTranslate(ByVal dx As Single, ByVal dy As Single, Order As MatrixOrder)
    GdipTranslatePathGradientTransform mHandle, dx, dy, Order
End Sub

Public Sub SetSclae(ByVal scaleX As Single, ByVal scaleY As Single, Order As MatrixOrder)
    GdipScalePathGradientTransform mHandle, scaleX, scaleY, Order
End Sub

Public Sub GetBlend(outBlend() As Long, outPos() As Single)
    Dim Count As Long
    
    Count = Me.BlendCount
    
    ReDim outBlend(Count - 1) As Long
    ReDim outPos(Count - 1) As Single
    
    GdipGetPathGradientBlend mHandle, outBlend(0), outPos(0), Count
End Sub

Public Sub SetBlend(Blend() As Long, Position() As Single)
    Dim Count As Long
    If (Not Blend) = -1 Or (Not Position) = -1 Then Exit Sub
    
    Count = UBound(Blend) - LBound(Blend) + 1
    GdipSetPathGradientBlend mHandle, Blend(LBound(Blend)), Position(LBound(Position)), Count
End Sub

Public Sub GetPresetBlend(outBlend() As Long, outPos() As Single)
    Dim Count As Long
    
    Count = Me.PresetBlendCount
    
    ReDim outBlend(Count - 1) As Long
    ReDim outPos(Count - 1) As Single
    
    GdipGetPathGradientPresetBlend mHandle, outBlend(0), outPos(0), Count
End Sub

Public Sub SetPresetBlend(Blend() As Long, Position() As Single)
    Dim Count As Long
    If (Not Blend) = -1 Or (Not Position) = -1 Then Exit Sub
    
    Count = UBound(Blend) - LBound(Blend) + 1
    GdipSetPathGradientPresetBlend mHandle, Blend(LBound(Blend)), Position(LBound(Position)), Count
End Sub

Public Sub SetBellShape(ByVal fFoucs As Single, ByVal fScale As Single)
    GdipSetPathGradientSigmaBlend mHandle, fFoucs, fScale
End Sub

Public Sub SetTriangularShape(ByVal fFoucs As Single, ByVal fScale As Single)
    GdipSetPathGradientLinearBlend mHandle, fFoucs, fScale
End Sub

Public Function GetBaseBrush() As cBrush
    Dim Brush As New cBrush
    Dim H As Long
    GdipCloneBrush mHandle, H
    Brush.Handle = H
    Set GetBaseBrush = Brush
End Function
