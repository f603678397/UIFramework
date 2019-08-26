VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Bitmap As cBitmap

Private Sub Class_Terminate(): Set Bitmap = Nothing: End Sub

Friend Sub Create(ByVal nWidth As Long, ByVal nHeight As Long)
    Dim b() As Byte
    Set Bitmap = New cBitmap
    Bitmap.CreateEmptyBitmap nWidth, nHeight
End Sub

Friend Function GetImage() As cImage
    Set GetImage = Bitmap.ToImage
End Function

Friend Function HandleMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef Result As Long) As Boolean
    HandleMessage = False
End Function


