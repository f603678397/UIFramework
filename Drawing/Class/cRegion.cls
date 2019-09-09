VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegion"
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

Friend Sub Create(): GdipCreateRegion mHandle: End Sub
Friend Sub Release(): GdipDeleteRegion mHandle: mHandle = 0: End Sub

Public Sub FromPath(Path As cPath)
    Release
    GdipCreateRegionPath Path.Handle, mHandle
End Sub

Public Sub FromRect(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    Dim mRect As RECTF
    With mRect
        .Left = X
        .Top = Y
        .Right = Width
        .Bottom = Height
    End With
    FromRect2 mRect
End Sub

Public Sub FromRect2(Rct As RECTF)
    Release
    GdipCreateRegionRect Rct, mHandle
End Sub

Public Function Clone() As cRegion
    Dim Rgn As cRegion
    Dim h As Long
    Rgn.Release
    GdipCloneRegion mHandle, h
    Rgn.Handle = h
    Set Clone = Rgn
End Function

Public Sub SetEmpty()
    GdipSetEmpty mHandle
End Sub

Public Function IsEmpty(Graphics As cGraphics) As Boolean
    Dim Result As Long
    GdipIsEmptyRegion mHandle, Graphics.Handle, Result
    IsEmpty = Result
End Function

Public Sub SetInfinite()
    GdipSetInfinite mHandle
End Sub

Public Function IsInfinite(Graphics As cGraphics) As Boolean
    Dim Result As Long
    GdipIsInfiniteRegion mHandle, Graphics.Handle, Result
    IsInfinite = Result
End Function

Public Function IsEqual(Rgn As cRegion, Graphics As cGraphics) As Boolean
    Dim Result As Long
    GdipIsEqualRegion mHandle, Rgn.Handle, Graphics.Handle, Result
    IsEqual = Result
End Function

Public Function IsPointVisible(ByVal X As Single, ByVal Y As Single, Graphics As cGraphics) As Boolean
    Dim Result As Long
    Dim h As Long
    
    If Not Graphics Is Nothing Then h = Graphics.Handle
    GdipIsVisibleRegionPoint mHandle, X, Y, h, Result
    IsPointVisible = Result
End Function

Public Function IsRectVisible(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Graphics As cGraphics) As Boolean
    Dim Result As Long
    Dim h As Long
    If Not Graphics Is Nothing Then h = Graphics.Handle
    GdipIsVisibleRegionRect mHandle, X, Y, Width, Height, h, Result
    IsRectVisible = Result
End Function

Public Sub CombineRect(Rct As RECTF, CombineMd As CombineMode)
    GdipCombineRegionRect mHandle, Rct, CombineMd
End Sub

Public Sub CombinePath(Path As cPath, CombineMd As CombineMode)
    GdipCombineRegionPath mHandle, Path.Handle, CombineMd
End Sub

Public Sub CombineRegion(Rgn As cRegion, CombineMd As CombineMode)
    GdipCombineRegionRegion mHandle, Rgn.Handle, CombineMd
End Sub

Public Sub SetTranslate(ByVal dx As Single, ByVal dy As Single)
    GdipTranslateRegion mHandle, dx, dy
End Sub

Public Sub SetTransform(Matrix As cMatrix)
    GdipTransformRegion mHandle, Matrix.Handle
End Sub

Public Function GetBounds(Graphics As cGraphics) As RECTF
    GdipGetRegionBounds mHandle, Graphics.Handle, GetBounds
End Function

Public Function GetScansCount(Matrix As cMatrix) As Long
    GdipGetRegionScansCount mHandle, GetScansCount, Matrix.Handle
End Function

Public Function GetScans(Matrix As cMatrix) As RECTF()
    Dim mRect() As RECTF
    Dim Count As Long
    
    Count = GetScansCount(Matrix)
    ReDim mRect(Count - 1) As RECTF
    GdipGetRegionScans mHandle, mRect(0), Count, Matrix.Handle
    GetScans = mRect
End Function
