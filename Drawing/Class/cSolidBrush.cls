VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSolidBrush"
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

Public Property Get Color() As Long: GdipGetSolidFillColor mHandle, Color: End Property
Public Property Let Color(ByVal nColor As Long): GdipSetSolidFillColor mHandle, nColor: End Property

Friend Sub Create()
    GdipCreateSolidFill ARGB(255, 0, 0, 0), mHandle
End Sub

Friend Sub Release()
    GdipDeleteBrush mHandle
    mHandle = 0
End Sub

Public Function Clone() As cSolidBrush
    Dim Brush As New cSolidBrush
    Dim H As Long
    Brush.Release
    GdipCloneBrush mHandle, H
    Brush.Handle = H
    Set Clone = Brush
End Function

Public Function GetBaseBrush() As cBrush
    Dim Brush As New cBrush
    Dim H As Long
    GdipCloneBrush mHandle, H
    Brush.Handle = H
    Set GetBaseBrush = Brush
End Function
