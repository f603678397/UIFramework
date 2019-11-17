VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPath"
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

Public Property Get BrushMode() As FillMode: GdipGetPathFillMode mHandle, BrushMode: End Property
Public Property Let BrushMode(emMode As FillMode): GdipSetPathFillMode mHandle, emMode: End Property

Friend Sub Create(): GdipCreatePath FillModeAlternate, mHandle: End Sub
Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeletePath mHandle
    mHandle = 0
End Sub

Public Function Clone() As cPath
    Dim Path As New cPath
    Dim H As Long
    Path.Release
    GdipClonePath mHandle, H
    Path.Handle = H
    Set Clone = Path
End Function

Public Sub StartFigure()
    GdipStartPathFigure mHandle
End Sub

Public Sub CloseFigure()
    GdipClosePathFigure mHandle
End Sub

Public Sub CloseAllFigures()
    GdipClosePathFigures mHandle
End Sub

Public Sub ReSet()
    GdipResetPath mHandle
End Sub

Public Sub Reverse()
    GdipReversePath mHandle
End Sub

Public Sub Flatten(Matrix As cMatrix, Flatness As Single)
    If Matrix Is Nothing Then
        GdipFlattenPath mHandle, 0, Flatness
    Else
        GdipFlattenPath mHandle, Matrix.Handle, Flatness
    End If
End Sub

Public Sub Widen(Pen As cPen, Matrix As cMatrix, Flatness As Single)
    If Matrix Is Nothing Then
        GdipWidenPath mHandle, Pen.Handle, 0, Flatness
    Else
        GdipWidenPath mHandle, Pen.Handle, Matrix.Handle, Flatness
    End If
End Sub

Public Sub OutLine(Matrix As cMatrix, Flatness As Single)
    If Matrix Is Nothing Then
        GdipWindingModeOutline mHandle, 0, Flatness
    Else
        GdipWindingModeOutline mHandle, Matrix.Handle, Flatness
    End If
End Sub

Public Sub Warp(Points() As POINTF, mRect As RECTF, Matrix As cMatrix, emWarp As WarpMode, Flatness As Single)
    Dim Count As Long
    If (Not Points) = -1 Then Exit Sub
    Count = UBound(Points) - LBound(Points) + 1
    GdipWarpPath mHandle, Matrix.Handle, Points(LBound(Points)), Count, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, emWarp, Flatness
End Sub


Public Sub SetMarker()
    GdipSetPathMarker mHandle
End Sub

Public Sub ClearMarker()
    GdipClearPathMarkers mHandle
End Sub

Public Sub Transform(Matrix As cMatrix)
    GdipTransformPath mHandle, Matrix.Handle
End Sub

Public Function GetPointCount() As Long
    GdipGetPointCount mHandle, GetPointCount
End Function

Public Function GetPoints() As POINTF()
    Dim Count As Long
    Dim P() As POINTF
    Count = GetPointCount
    ReDim P(Count - 1) As POINTF
    GdipGetPathPoints mHandle, P(0), Count
    GetPoints = P
End Function

Public Function GetLastPoint() As POINTF
    GdipGetPathLastPoint mHandle, GetLastPoint
End Function

Public Function GetBounds(Matrix As cMatrix, Pen As cPath) As RECTF
    Dim M As Long, P As Long
    If Not Matrix Is Nothing Then M = Matrix.Handle
    If Not Pen Is Nothing Then P = Pen.Handle
    GdipGetPathWorldBounds mHandle, GetBounds, M, P
End Function

Public Function GetData() As PathData
    GdipGetPathData mHandle, GetData
End Function

Public Function IsPointVisible(ByVal X As Single, ByVal Y As Single, Graphics As cGraphics) As Boolean
    Dim Result As Long
    Dim G As Long
    If Not Graphics Is Nothing Then G = Graphics.Handle
    GdipIsVisiblePathPoint mHandle, X, Y, G, Result
    IsPointVisible = Result
End Function

Public Function IsOutLineVisible(Pen As cPen, ByVal X As Single, ByVal Y As Single, Graphics As cGraphics)
    Dim Result As Long
    Dim G As Long
    If Not Graphics Is Nothing Then G = Graphics.Handle
    GdipIsOutlineVisiblePathPoint mHandle, X, Y, Pen.Handle, G, Result
    IsOutLineVisible = Result
End Function

Public Sub AddLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    GdipAddPathLine mHandle, x1, y1, x2, y2
End Sub

Public Sub AddLines(Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipAddPathLine2 mHandle, Ps(LBound(Ps)), Count
End Sub

Public Sub AddArc(ByVal X As Single, ByVal Y As Single, ByVal w As Single, ByVal H As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    GdipAddPathArc mHandle, X, Y, w, H, startAngle, sweepAngle
End Sub

Public Sub AddBezier(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single)
    GdipAddPathBezier mHandle, x1, y1, x2, y2, x3, y3, x4, y4
End Sub

Public Sub AddBeziers(Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipAddPathBeziers mHandle, Ps(LBound(Ps)), Count
End Sub

Public Sub AddCurve(Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipAddPathCurve mHandle, Ps(LBound(Ps)), Count
End Sub

Public Sub AddCurve2(Ps() As POINTF, ByVal tension As Single)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipAddPathCurve2 mHandle, Ps(LBound(Ps)), Count, tension
End Sub

Public Sub AddCurve3(Ps() As POINTF, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipAddPathCurve3 mHandle, Ps(LBound(Ps)), Count, offset, numberOfSegments, tension
End Sub

Public Sub AddClosedCurve(Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipAddPathClosedCurve mHandle, Ps(LBound(Ps)), Count
End Sub

Public Sub AddClosedCurve2(Ps() As POINTF, ByVal tension As Single)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipAddPathClosedCurve2 mHandle, Ps(LBound(Ps)), Count, tension
End Sub

Public Sub AddEllipse(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    GdipAddPathEllipse mHandle, X, Y, Width, Height
End Sub

Public Sub AddPath(Path As cPath, ByVal bConnect As Boolean)
    GdipAddPathPath mHandle, Path.Handle, IIf(bConnect, 1, 0)
End Sub

Public Sub AddPie(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    GdipAddPathPie mHandle, X, Y, Width, Height, startAngle, sweepAngle
End Sub

Public Sub AddPolygon(Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipAddPathPolygon mHandle, Ps(LBound(Ps)), Count
End Sub

Public Sub AddRectangle(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    GdipAddPathRectangle mHandle, X, Y, Width, Height
End Sub

Public Sub AddRectangles(Rects() As RECTF)
    Dim Count As Long
    If (Not Rects) = -1 Then Exit Sub
    Count = UBound(Rects) - LBound(Rects) + 1
    GdipAddPathRectangles mHandle, Rects(LBound(Rects)), Count
End Sub

Public Sub AddString(ByVal Text As String, Fam As cFontFamily, Style As FontStyle, ByVal Size As Single, Layout As RECTF, strFormat As cStringFormat)
    GdipAddPathString mHandle, StrPtr(Text), -1, Fam.Handle, Style, Size, Layout, strFormat.Handle
End Sub
