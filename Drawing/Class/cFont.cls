VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long): mHandle = nHandle: End Property

Friend Sub Release()
    GdipDeleteFont mHandle
    mHandle = 0
End Sub

Public Sub FromHDC(ByVal hdc As Long)
    Release
    GdipCreateFontFromDC hdc, mHandle
End Sub

Public Sub FromFamily(Fam As cFontFamily, ByVal fSize As Single, nStyle As FontStyle, Unit As GpUnit)
    Release
    GdipCreateFont Fam.Handle, fSize, nStyle, Unit, mHandle
End Sub

Public Function Clone() As cFont
    Dim mFont As New cFont
    GdipCloneFont mHandle, mFont.Handle
    Set Clone = mFont
End Function

Public Function GetFamily() As cFontFamily
    Dim Fam As New cFontFamily
    GdipGetFamily mHandle, Fam.Handle
    Set GetFamily = Fam
End Function

Public Function GetSize() As Single
    GdipGetFontSize mHandle, GetSize
End Function

Public Function GetStyle() As FontStyle
    GdipGetFontStyle mHandle, GetStyle
End Function

Public Function GetUnit() As GpUnit
    GdipGetFontUnit mHandle, GetUnit
End Function

Public Function GetHeight(Graphics As cGraphics) As Single
    GdipGetFontHeight mHandle, Graphics.Handle, GetHeight
End Function

Public Function GetHeightGivenDPI(ByVal DPI As Single)
    GdipGetFontHeightGivenDPI mHandle, DPI, GetHeightGivenDPI
End Function