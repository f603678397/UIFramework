VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWidgetManager"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Layout              As cLayout

Public Enum PresetThemeEnum
    DrakTheme
    LightTheme
End Enum

Public Enum ViewMessageEnum
    VM_EMPTY = 0
    VM_OPTION = 1
End Enum

Private Sub Class_Terminate(): Release: End Sub

Private Sub Release()
    Set Layout = Nothing
    Set ViewList = Nothing
End Sub

Private Function GetBackColor() As Long
    Dim Color As Long
    Select Case Preset
        Case PresetThemeEnum.DrakTheme
            Color = cColor.FromARGB(255, 50, 50, 50)
        Case PresetThemeEnum.LightTheme
            Color = cColor.FromARGB(255, 225, 225, 225)
    End Select
    GetBackColor = Color
End Function

Public Sub BindLayout(dstLayout As cLayout)
    Set Layout = dstLayout
    
    Layout.BackgroundColor = GetBackColor
End Sub

Public Sub SetPresetTheme(ByVal nFlag As PresetThemeEnum)
    Dim View As Object
    Dim i As Long
    
    Preset = nFlag
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            View.SetTheme nFlag
        Set View = Nothing
    Next
    
    If Layout Is Nothing Then Exit Sub
    Layout.BackgroundColor = GetBackColor
    Layout.ReDrawAllView
End Sub

Public Function CreateButton(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cButton
    Dim Button As New cButton
    
    If Parent Is Nothing Then
        Button.Create Layout, "Button", L, T, W, H
    Else
        Button.Create Parent, "Button", L, T, W, H
    End If
    
    ViewList.Add Button
    Set CreateButton = Button
End Function

Public Function CreateLabel(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cLabel
    Dim Label As New cLabel
    
    If Parent Is Nothing Then
        Label.Create Layout, "Label", L, T, W, H
    Else
        Label.Create Parent, "Label", L, T, W, H
    End If
    
    ViewList.Add Label
    Set CreateLabel = Label
End Function

Public Function CreateOption(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cOption
    Dim Option1 As New cOption
    
    If Parent Is Nothing Then
        Option1.Create Layout, "Option", L, T, W, H
    Else
        Option1.Create Parent, "Option", L, T, W, H
    End If
    
    ViewList.Add Option1
    Set CreateOption = Option1
End Function
