VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPen"
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

Public Property Get Color() As Long: GdipGetPenColor mHandle, Color: End Property
Public Property Let Color(ByVal nColor As Long): GdipSetPenColor mHandle, nColor: End Property

Public Property Get Width() As Single: GdipGetPenWidth mHandle, Width: End Property
Public Property Let Width(ByVal fWidth As Single): GdipSetPenWidth mHandle, fWidth: End Property

Public Property Get Unit() As GpUnit: GdipGetPenUnit mHandle, Unit: End Property
Public Property Let Unit(ByRef emUnit As GpUnit): GdipSetPenUnit mHandle, emUnit: End Property

Public Property Get FillType() As PenType: GdipGetPenFillType mHandle, FillType: End Property

Public Property Get Alignment() As PenAlignment: GdipGetPenMode mHandle, Alignment: End Property
Public Property Let Alignment(ByRef emMode As PenAlignment): GdipSetPenMode mHandle, emMode: End Property

Public Property Get LineJoin() As GpLineJoin: GdipGetPenLineJoin mHandle, LineJoin: End Property
Public Property Let LineJoin(ByRef emJoin As GpLineJoin): GdipSetPenLineJoin mHandle, emJoin: End Property

Public Property Get StartCap() As LineCap: GdipGetPenStartCap mHandle, StartCap: End Property
Public Property Let StartCap(ByRef emCap As LineCap): GdipSetPenStartCap mHandle, emCap: End Property

Public Property Get EndCap() As LineCap: GdipGetPenEndCap mHandle, EndCap: End Property
Public Property Let EndCap(ByRef emCap As LineCap): GdipSetPenEndCap mHandle, emCap: End Property

Public Property Get MiterLimit() As Single: GdipGetPenMiterLimit mHandle, MiterLimit: End Property
Public Property Let MiterLimit(ByVal fLimit As Single): GdipSetPenMiterLimit mHandle, fLimit: End Property

Friend Sub Create(): GdipCreatePen1 ARGB(255, 0, 0, 0), 1, GlobalUnit, mHandle: End Sub
Friend Sub Release(): GdipDeletePen mHandle: mHandle = 0: End Sub

Public Function Clone() As cPen
    Dim Pen As New cPen
    Dim H As Long
    Pen.Release
    GdipClonePen mHandle, H
    Pen.Handle = H
    Set Clone = Pen
End Function

Public Sub FromBrush(Brush As cBrush)
    Release
    GdipCreatePen2 Brush.Handle, 1, GlobalUnit, mHandle
End Sub

Public Sub SetBrush(Brush As cBrush)
    GdipSetPenBrushFill mHandle, Brush.Handle
End Sub

Public Function GetBrush() As cBrush
    Dim Brush As New cBrush
    Dim H As Long
    GdipGetPenBrushFill mHandle, H
    Brush.Handle = H
    Set GetBrush = Brush
End Function

Public Function GetDashOffset() As Single: GdipGetPenDashOffset mHandle, GetDashOffset: End Function
Public Sub SetDashOffset(ByVal fOffset As Single): GdipSetPenDashOffset mHandle, fOffset: End Sub

Public Function GetDashCap() As DashCap: GdipGetPenDashCap mHandle, GetDashCap: End Function
Public Sub SetDashCap(emCap As DashCap): GdipSetPenDashCap mHandle, emCap: End Sub

Public Function GetDashStyle() As DashStyle: GdipGetPenDashStyle mHandle, GetDashStyle: End Function
Public Sub SetDashStyle(emStyle As DashStyle): GdipSetPenDashStyle mHandle, emStyle: End Sub

Public Function GetDashCount() As Long: GdipGetPenDashCount mHandle, GetDashCount: End Function
Public Function GetCompoundCount() As Long: GdipGetPenCompoundCount mHandle, GetCompoundCount: End Function

Public Function GetCompoundArray() As Single()
    Dim n As Long, dash() As Single
    n = GetCompoundCount
    ReDim dash(n - 1) As Single
    GdipGetPenCompoundArray mHandle, dash(0), n
    GetCompoundArray = dash
End Function

Public Sub SetCompoundArray(dash() As Single)
    Dim Count As Long
    If (Not dash) = -1 Then Exit Sub
    Count = UBound(dash) - LBound(dash) + 1
    GdipSetPenCompoundArray mHandle, dash(LBound(dash)), Count
End Sub

Public Function GetDashArray() As Single()
    Dim n As Long, dash() As Single
    n = GetDashCount
    ReDim dash(n - 1) As Single
    GdipGetPenDashArray mHandle, dash(0), n
    GetDashArray = dash
End Function

Public Sub SetDashArray(dash() As Single)
    Dim Count As Long
    If (Not dash) = -1 Then Exit Sub
    Count = UBound(dash) - LBound(dash) + 1
    GdipSetPenDashArray mHandle, dash(LBound(dash)), Count
End Sub

Public Function GetCustomStartCap() As cCustomLineCap
    Dim Cap As New cCustomLineCap
    Dim H As Long
    Cap.Release
    GdipGetPenCustomStartCap mHandle, H
    Cap.Handle = H
    Set GetCustomStartCap = Cap
End Function

Public Sub SetCustomStartCap(Cap As cCustomLineCap)
    GdipSetPenCustomStartCap mHandle, Cap.Handle
End Sub

Public Function GetCustomEndCap() As cCustomLineCap
    Dim Cap As New cCustomLineCap
    Dim H As Long
    Cap.Release
    GdipGetPenCustomEndCap mHandle, H
    Cap.Handle = H
    Set GetCustomEndCap = Cap
End Function

Public Sub SetCustomEndCap(Cap As cCustomLineCap)
    GdipSetPenCustomEndCap mHandle, Cap.Handle
End Sub

Public Sub SetTransform(Matrix As cMatrix)
    GdipSetPenTransform mHandle, Matrix.Handle
End Sub

Public Function GetTransform() As cMatrix
    Dim Matrix As New cMatrix
    Dim H As Long
    Matrix.Release
    GdipGetPenTransform mHandle, H
    Matrix.Handle = H
    Set GetTransform = Matrix
End Function

Public Sub SetRotate(ByVal Angle As Single, Order As MatrixOrder): GdipRotatePenTransform mHandle, Angle, Order: End Sub
Public Sub SetScale(ByVal sx As Single, ByVal sy As Single, Order As MatrixOrder): GdipScalePenTransform mHandle, sx, sy, Order: End Sub
Public Sub SetTranslate(ByVal dx As Single, ByVal dy As Single, Order As MatrixOrder): GdipTranslatePenTransform mHandle, dx, dy, Order: End Sub
Public Sub ResetTransform(): GdipResetPenTransform mHandle: End Sub
