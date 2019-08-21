VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCustomLineCap"
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

Public Property Get BaseCap() As LineCap: GdipGetCustomLineCapBaseCap mHandle, BaseCap: End Property
Public Property Let BaseCap(emCap As LineCap): GdipSetCustomLineCapBaseCap mHandle, emCap: End Property

Public Property Get BaseInset() As Single: GdipGetCustomLineCapBaseInset mHandle, BaseInset: End Property
Public Property Let BaseInset(ByVal fInset As Single): GdipSetCustomLineCapBaseInset mHandle, fInset: End Property

Public Property Get WidthScale() As Single: GdipGetCustomLineCapWidthScale mHandle, WidthScale: End Property
Public Property Let WidthSclae(ByVal fScale As Single): GdipSetCustomLineCapWidthScale mHandle, fScale: End Property

Friend Sub Create()
    GdipCreateCustomLineCap 0, 0, LineCapCustom, 1, mHandle
End Sub

Friend Sub Release()
    GdipDeleteCustomLineCap mHandle
    mHandle = 0
End Sub

Public Function Clone() As cCustomLineCap
    Dim Cap As New cCustomLineCap
    Cap.Release
    GdipCloneCustomLineCap mHandle, Cap.Handle
    Set Clone = Cap
End Function

Public Sub GetStrokeCaps(ByRef outStart As LineCap, ByRef outEnd As LineCap)
    GdipGetCustomLineCapStrokeCaps mHandle, outStart, outEnd
End Sub

Public Sub SetStrokeCaps(StartCap As LineCap, EndCap As LineCap)
    GdipSetCustomLineCapStrokeCaps mHandle, StartCap, EndCap
End Sub

Public Function GetStrokeJoin() As GpLineJoin
    GdipGetCustomLineCapStrokeJoin mHandle, GetStrokeJoin
End Function

Public Sub SetStrokeJoin(emJoin As GpLineJoin)
    GdipSetCustomLineCapStrokeJoin mHandle, emJoin
End Sub