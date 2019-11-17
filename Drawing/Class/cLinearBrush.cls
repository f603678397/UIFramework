VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLinearBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

Private mColor1 As Long
Private mColor2 As Long
Private P1 As POINTF, P2 As POINTF
Private mWrap As WrapMode

Private Sub Class_Initialize()
    mColor1 = ARGB(255, 0, 0, 0)
    mColor2 = ARGB(255, 255, 255, 255)
    mWrap = WrapModeTile
    Create
End Sub
Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long)
    Dim mRect As RECTF
    mHandle = nHandle
    mRect = GetRect()
    P1.X = mRect.Left
    P1.Y = mRect.Top
    P2.X = mRect.Left + mRect.Right
    P2.Y = mRect.Top + mRect.Bottom
    mColor1 = StartColor
    mColor2 = EndColor
    mWrap = FillWrapMode
End Property

Public Property Get StartColor() As Long
    Dim Color(0 To 1) As Long
    GdipGetLineColors mHandle, Color(0)
    StartColor = Color(0)
End Property

Public Property Let StartColor(ByVal nColor As Long)
    mColor1 = nColor
    GdipSetLineColors mHandle, mColor1, mColor2
End Property

Public Property Get EndColor() As Long
    Dim Color(0 To 1) As Long
    GdipGetLineColors mHandle, Color(0)
    EndColor = Color(1)
End Property

Public Property Let EndColor(ByVal nColor As Long)
    mColor2 = nColor
    GdipSetLineColors mHandle, mColor1, mColor2
End Property

Public Property Get FillWrapMode() As WrapMode: GdipGetLineWrapMode mHandle, FillWrapMode: End Property
Public Property Let FillWrapMode(ByVal emMode As WrapMode): mWrap = emMode: GdipSetLineWrapMode mHandle, emMode: End Property

Public Property Get GammaCorrection() As Boolean
    Dim n As Long
    GdipGetLineGammaCorrection mHandle, n
    GammaCorrection = n
End Property

Public Property Let GammaCorrection(ByVal bGamma As Boolean)
    GdipSetLineGammaCorrection mHandle, IIf(bGamma, 1, 0)
End Property

Friend Sub Create()
    GdipCreateLineBrush P1, P2, mColor1, mColor2, mWrap, mHandle
End Sub

Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeleteBrush mHandle
    mHandle = 0
End Sub

Public Function Clone() As cLinearBrush
    Dim Brush As New cLinearBrush
    Dim H As Long
    Brush.Release
    GdipCloneBrush mHandle, H
    Brush.Handle = H
    Set Clone = Brush
End Function

Public Function GetRect() As RECTF
    Dim mRect As RECTF
    GdipGetLineRect mHandle, mRect
    GetRect = mRect
End Function

Public Sub SetStartPoint(ByVal X As Single, ByVal Y As Single)
    P1.X = X: P1.Y = Y
    Release
    Create
End Sub

Public Sub SetEndPoint(ByVal X As Single, ByVal Y As Single)
    P2.X = X: P2.Y = Y
    Release
    Create
End Sub

Public Sub SetColors(ByVal nColor1 As Long, ByVal nColor2 As Long)
    mColor1 = nColor1
    mColor2 = nColor2
    GdipSetLineColors mHandle, mColor1, mColor2
End Sub

Public Function GetBlendCount() As Long
    GdipGetLineBlendCount mHandle, GetBlendCount
End Function

Public Function GetPresetBlendCount() As Long
    GdipGetLinePresetBlendCount mHandle, GetPresetBlendCount
End Function

Public Sub GetBlend(OutColor() As Long, outPos() As Single)
    Dim Count As Long
    Dim C() As Long, P() As Single
    Count = GetBlendCount
    
    ReDim C(Count - 1) As Long
    ReDim P(Count - 1) As Single
    
    GdipGetLineBlend mHandle, C(0), P(0), Count
    OutColor = C: outPos = P
End Sub

Public Sub SetBlend(Blend() As Long, Position() As Single)
    Dim Count As Long
    If (Not Blend) = -1 Or (Not Position) = -1 Then Exit Sub
    Count = UBound(Blend) - LBound(Blend) + 1
    GdipSetLineBlend mHandle, Blend(LBound(Blend)), Position(LBound(Position)), Count
End Sub

Public Sub GetPresetBlend(OutColor() As Long, outPos() As Single)
    Dim Count As Long
    Dim C() As Long, P() As Single
    Count = GetPresetBlendCount
    
    ReDim C(Count - 1) As Long
    ReDim P(Count - 1) As Single
    
    GdipGetLinePresetBlend mHandle, C(0), P(0), Count
    OutColor = C: outPos = P
End Sub

Public Sub SetPresetBlend(Blend() As Long, Position() As Single)
    Dim Count As Long
    If (Not Blend) = -1 Or (Not Position) = -1 Then Exit Sub
    Count = UBound(Blend) - LBound(Blend) + 1
    GdipSetLinePresetBlend mHandle, Blend(LBound(Blend)), Position(LBound(Position)), Count
End Sub

Public Sub SetBellShape(ByVal fFocus As Single, Optional ByVal fScale As Single = 1#)
    GdipSetLineSigmaBlend mHandle, fFocus, fScale
End Sub

Public Sub SetTriangularShape(ByVal fFocus As Single, Optional ByVal fScale As Single = 1#)
    GdipSetLineLinearBlend mHandle, fFocus, fScale
End Sub

Public Sub SetScale(ByVal sx As Single, ByVal sy As Single, Order As MatrixOrder)
    GdipScaleLineTransform mHandle, sx, sy, Order
End Sub

Public Sub SetRotate(ByVal Angle As Single, Order As MatrixOrder)
    GdipRotateLineTransform mHandle, Angle, Order
End Sub

Public Sub SetTranslate(ByVal dx As Single, ByVal dy As Single, Order As MatrixOrder)
    GdipTranslateLineTransform mHandle, dx, dy, Order
End Sub

Public Sub SetTransform(Matrix As cMatrix)
    GdipSetLineTransform mHandle, Matrix.Handle
End Sub

Public Function GetTransform() As cMatrix
    Dim Matrix As New cMatrix
    Dim H As Long
    Matrix.Release
    GdipGetLineTransform mHandle, H
    Matrix.Handle = H
    Set GetTransform = Matrix
End Function

Public Sub ResetTransform()
    GdipResetLineTransform mHandle
End Sub

Public Function GetBaseBrush() As cBrush
    Dim Brush As New cBrush
    Dim H As Long
    GdipCloneBrush mHandle, H
    Brush.Handle = H
    Set GetBaseBrush = Brush
End Function
