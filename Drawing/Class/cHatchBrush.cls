VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHatchBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mHandle As Long

Private mForeColor As Long
Private mBackColor As Long
Private mStyle As HatchStyle

Private Sub Class_Initialize(): Create: End Sub
Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long)
    mHandle = nHandle
    mStyle = Me.Style
    mForeColor = Me.ForeColor
    mBackColor = Me.BackColor
End Property

Public Property Get Style() As HatchStyle
    GdipGetHatchStyle mHandle, Style
End Property

Public Property Let Style(emStyle As HatchStyle)
    Release
    mStyle = emStyle
    GdipCreateHatchBrush mStyle, mForeColor, mBackColor, mHandle
End Property

Public Property Get ForeColor() As Long
    GdipGetHatchForegroundColor mHandle, ForeColor
End Property

Public Property Let ForColor(ByVal nColor As Long)
    Release
    mForeColor = nColor
    GdipCreateHatchBrush mStyle, mForeColor, mBackColor, mHandle
End Property

Public Property Get BackColor() As Long
    GdipGetHatchBackgroundColor mHandle, BackColor
End Property

Public Property Let BackColor(ByVal nColor As Long)
    Release
    mBackColor = nColor
    GdipCreateHatchBrush mStyle, mForeColor, mBackColor, mHandle
End Property

Friend Sub Create()
    mForeColor = ARGB(255, 0, 0, 0)
    mBackColor = ARGB(255, 255, 255, 255)
    mStyle = HatchStyleDarkDownwardDiagonal
    GdipCreateHatchBrush mStyle, mForeColor, mBackColor, mHandle
End Sub

Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeleteBrush mHandle
    mHandle = 0
End Sub

Public Function Clone() As cHatchBrush
    Dim Brush As New cHatchBrush
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
