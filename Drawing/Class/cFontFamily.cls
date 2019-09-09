VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFontFamily"
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
    GdipDeleteFontFamily mHandle
    mHandle = 0
End Sub

Public Sub FromName(ByVal sName As String)
    GdipCreateFontFamilyFromName sName, 0, mHandle
End Sub

Public Function Clone() As cFontFamily
    Dim Fam As New cFontFamily
    Dim h As Long
    GdipCloneFontFamily mHandle, h
    Fam.Handle = h
    Set Clone = Fam
End Function

Public Sub GenericSansSerif()
    GdipGetGenericFontFamilySansSerif mHandle
End Sub

Public Sub GenericSerif()
    GdipGetGenericFontFamilySerif mHandle
End Sub

Public Sub GenericMonospace()
    GdipGetGenericFontFamilyMonospace mHandle
End Sub

Public Function GetFamilyName() As String
    Dim sName As String
    sName = Space$(32)
    GdipGetFamilyName mHandle, StrPtr(sName), 0
    
    sName = Replace$(sName, Chr(0), "")
    GetFamilyName = sName
End Function

Public Function IsStyleAvailable(nStyle As FontStyle) As Boolean
    Dim Result As Long
    GdipIsStyleAvailable mHandle, nStyle, Result
    IsStyleAvailable = Result
End Function

Public Function GetEmHeight(nStyle As FontStyle) As Integer
    GdipGetEmHeight mHandle, nStyle, GetEmHeight
End Function

Public Function GetCellAscent(nStyle As FontStyle) As Integer
    GdipGetCellAscent mHandle, nStyle, GetCellAscent
End Function

Public Function GetCellDescent(nStyle As FontStyle) As Integer
    GdipGetCellDescent mHandle, nStyle, GetCellDescent
End Function

Public Function GetLineSpacing(nStyle As FontStyle) As Integer
    GdipGetLineSpacing mHandle, nStyle, GetLineSpacing
End Function
