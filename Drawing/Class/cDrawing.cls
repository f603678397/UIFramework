VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrawing"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Init()
    If Token <> 0 Then Exit Sub
    
    Dim uIn As GdiplusStartupInput
    
    uIn.GdiplusVersion = 1
    GdiplusStartup Token, uIn
    GlobalUnit = UnitPixel
End Sub

Public Sub Shutdown()
    GdiplusShutdown Token
    Token = 0
End Sub

Public Sub Alloc(ByVal Size As Long)
    GdipAlloc Size
End Sub

Public Sub Free(ByVal ptr As Long)
    GdipFree ptr
End Sub

Public Sub SetUnit(ByVal nUnit As GpUnit)
    GlobalUnit = nUnit
End Sub

Public Function GetUnit() As GpUnit
    GetUnit = GlobalUnit
End Function

Public Function CreateGraphicsFromHWND(ByVal Hwnd As Long) As cGraphics
    Dim Graphics As New cGraphics
    Dim H As Long
    
    GdipCreateFromHWND Hwnd, H
    Graphics.Handle = H
    
    Set CreateGraphicsFromHWND = Graphics
End Function

Public Function CreateGraphicsFromHDC(ByVal hdc As Long) As cGraphics
    Dim Graphics As New cGraphics
    Dim H As Long
    
    GdipCreateFromHDC hdc, H
    Graphics.Handle = H
    Set CreateGraphicsFromHDC = Graphics
End Function

Public Function NewRectF(ByVal nLeft As Single, ByVal nTop As Single, ByVal nRight As Single, ByVal nBottom As Single) As Drawing2D.RECTF
    With NewRectF
        .Left = nLeft
        .Top = nTop
        .Right = nRight
        .Bottom = nBottom
    End With
End Function

Public Function NewPointF(ByVal X As Single, ByVal Y As Single) As Drawing2D.POINTF
    With NewPointF
        .X = X
        .Y = Y
    End With
End Function

Public Function NewPointFArray(ParamArray xy()) As Drawing2D.POINTF()
    Dim Count As Long
    Dim pts() As POINTF
    Count = Int((UBound(xy) + 1) / 2)
    
    If Count < 1 Then Exit Function
    
    ReDim pts(Count - 1) As POINTF
    
    Dim i As Long
    
    For i = 0 To Count - 1
        pts(i).X = xy(i * 2)
        pts(i).Y = xy(i * 2 + 1)
    Next
    
    NewPointFArray = pts
End Function

Public Function NewSolidBrush(ByVal nColor As Long) As cSolidBrush
    Dim Brush As New cSolidBrush
    Brush.Color = nColor
    Set NewSolidBrush = Brush
End Function

Public Function NewPen(ByVal nColor As Long, ByVal Width As Single) As cPen
    Dim Pen As cPen
    Pen.Color = nColor
    Pen.Width = Width
    Set NewPen = Pen
End Function
