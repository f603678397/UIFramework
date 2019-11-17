VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStringFormat"
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

Friend Sub Create()
    GdipCreateStringFormat 0, 0, mHandle
End Sub

Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeleteStringFormat mHandle
    mHandle = 0
End Sub

Public Sub GenericDefault()
    Release
    GdipStringFormatGetGenericDefault mHandle
End Sub

Public Sub GenericTypographic()
    Release
    GdipStringFormatGetGenericTypographic mHandle
End Sub

Public Function Clone() As cStringFormat
    Dim sf As New cStringFormat
    Dim H As Long
    sf.Release
    GdipCloneStringFormat mHandle, H
    sf.Handle = H
    Set Clone = sf
End Function

Public Property Get Flags() As StringFormatFlags
    GdipGetStringFormatFlags mHandle, Flags
End Property

Public Property Let Flags(ByVal nFlags As StringFormatFlags)
    GdipSetStringFormatFlags mHandle, nFlags
End Property

Public Property Get Align() As StringAlignment
    GdipGetStringFormatAlign mHandle, Align
End Property

Public Property Let Align(ByVal nAlign As StringAlignment)
    GdipSetStringFormatAlign mHandle, nAlign
End Property

Public Property Get LineAlign() As StringAlignment
    GdipGetStringFormatLineAlign mHandle, LineAlign
End Property

Public Property Let LineAlign(ByVal nAlign As StringAlignment)
    GdipSetStringFormatLineAlign mHandle, nAlign
End Property

Public Property Get Trimming() As StringTrimming
    GdipGetStringFormatTrimming mHandle, Trimming
End Property

Public Property Let Trimming(ByVal nTrim As StringTrimming)
    GdipSetStringFormatTrimming mHandle, nTrim
End Property

Public Property Get HotkeyPrefix() As GpHotkeyPrefix
    GdipGetStringFormatHotkeyPrefix mHandle, HotkeyPrefix
End Property

Public Property Let HotkeyPrefix(ByVal nPrefix As GpHotkeyPrefix)
    GdipSetStringFormatHotkeyPrefix mHandle, nPrefix
End Property

Public Property Get DigitSubstitution() As StringDigitSubstitute
    GdipGetStringFormatDigitSubstitution mHandle, 0, DigitSubstitution
End Property

Public Property Let DigitSubstitution(ByVal n As StringDigitSubstitute)
    GdipSetStringFormatDigitSubstitution mHandle, 0, n
End Property

Public Sub SetTabStops(ByVal firstOffset As Single, tabStops() As Single)
    Dim Count As Long
    If (Not tabStops) = -1 Then Exit Sub
    Count = UBound(tabStops) - LBound(tabStops) + 1
    GdipSetStringFormatTabStops mHandle, firstOffset, Count, tabStops(LBound(tabStops))
End Sub

Public Function GetTabStopCount() As Long
    GdipGetStringFormatTabStopCount mHandle, GetTabStopCount
End Function

Public Function GetTabStops(Optional outOffset As Single) As Single()
    Dim Count As Long
    Dim Stops() As Single
    Count = GetTabStopCount
    ReDim Stops(Count - 1) As Single
    GdipGetStringFormatTabStops mHandle, Count, outOffset, Stops(0)
    GetTabStops = Stops
End Function

Public Function GetMeasurableCharacterRangeCount() As Long
    GdipGetStringFormatMeasurableCharacterRangeCount mHandle, GetMeasurableCharacterRangeCount
End Function

Public Sub SetMeasurableCharacterRanges(ranges() As CharacterRange)
    Dim Count As Long
    If (Not ranges) = -1 Then Exit Sub
    Count = UBound(ranges) - LBound(ranges) + 1
    GdipSetStringFormatMeasurableCharacterRanges mHandle, Count, ranges(LBound(ranges))
End Sub
