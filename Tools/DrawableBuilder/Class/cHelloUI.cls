VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHelloUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cBaseUI

Private View        As cView

Public Sub Create(Parent As Object, ByVal W As Single, ByVal H As Single)
    Set View = Parent.CreateView(0, 0, W, H)
End Sub

Private Function cBaseUI_GetView() As Framework.cView
    Set cBaseUI_GetView = View
End Function

Private Sub cBaseUI_Move(ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer)
    View.Move Left, Top, Width, Height
End Sub
