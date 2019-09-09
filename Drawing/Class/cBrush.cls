VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long): mHandle = nHandle: End Property

Friend Sub Release()
    GdipDeleteBrush mHandle
    mHandle = 0
End Sub

Public Function GetType() As BrushType
    GdipGetBrushType mHandle, GetType
End Function

Public Function Clone() As cBrush
    Dim Brush As New cBrush
    Dim h As Long
    GdipCloneBrush mHandle, h
    Brush.Handle = h
    Set Clone = Brush
End Function

Public Function ToSolidBrush() As cSolidBrush
    Dim Brush As New cSolidBrush
    Brush.Release
    Brush.Handle = mHandle
    Set ToSolidBrush = Brush
End Function

Public Function ToLinearBrush() As cLinearBrush
    Dim Brush As New cLinearBrush
    Brush.Release
    Brush.Handle = mHandle
    Set ToLinearBrush = Brush
End Function

Public Function ToHatchBrush() As cHatchBrush
    Dim Brush As New cHatchBrush
    Brush.Release
    Brush.Handle = mHandle
    Set ToHatchBrush = Brush
End Function

Public Function ToPathBrush() As cPathBrush
    Dim Brush As New cPathBrush
    Brush.Handle = mHandle
    Set ToHatchBrush = Brush
End Function

Public Function ToTextureBrush() As cTextureBrush
    Dim Brush As New cTextureBrush
    Brush.Handle = mHandle
    Set ToTextureBrush = Brush
End Function
