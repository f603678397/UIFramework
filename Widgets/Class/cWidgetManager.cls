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
            Color = cColor.FromARGB(255, 45, 45, 45)
        Case PresetThemeEnum.LightTheme
            Color = cColor.FromARGB(255, 235, 235, 235)
    End Select
    GetBackColor = Color
End Function

Public Sub BindLayout(dstLayout As cLayout)
    Set Layout = dstLayout
    If dstLayout Is Nothing Then Exit Sub
    
    Layout.BackgroundColor = GetBackColor
    SetClassLongA Layout.GetActivity.hWnd, GCL_HCURSOR, ByVal 0&
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

Public Function CreateFrame(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cFrame
    Dim Frame As New cFrame
    
    If Parent Is Nothing Then
        Frame.Create Layout, "Frame", L, T, W, H
    Else
        Frame.Create Parent, "Frame", L, T, W, H
    End If
    
    ViewList.Add Frame
    Set CreateFrame = Frame
End Function

Public Function CreateProgressBar(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cProgressBar
    Dim Progress As New cProgressBar
    
    If Parent Is Nothing Then
        Progress.Create Layout, "ProgressBar", L, T, W, H
    Else
        Progress.Create Parent, "ProgressBar", L, T, W, H
    End If
    
    ViewList.Add Progress
    Set CreateProgressBar = Progress
End Function

Public Function CreateWaiting(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cWaiting
    Dim Waiting As New cWaiting
    If Parent Is Nothing Then
        Waiting.Create Layout, "Waiting", L, T, W, H
    Else
        Waiting.Create Parent, "Waiting", L, T, W, H
    End If
    ViewList.Add Waiting
    Set CreateWaiting = Waiting
End Function

Public Function CreateCheckBox(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cCheckBox
    Dim CheckBox As New cCheckBox
    If Parent Is Nothing Then
        CheckBox.Create Layout, "CheckBox", L, T, W, H
    Else
        CheckBox.Create Parent, "CheckBox", L, T, W, H
    End If
    ViewList.Add CheckBox
    Set CreateCheckBox = CheckBox
End Function

Public Function CreateImageView(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cImageView
    Dim Img As New cImageView
    If Parent Is Nothing Then
        Img.Create Layout, "ImageView", L, T, W, H
    Else
        Img.Create Parent, "ImageView", L, T, W, H
    End If
    ViewList.Add Img
    Set CreateImageView = Img
End Function

Public Function CreateVScrollBar(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cVScrollBar
    Dim VScroll As New cVScrollBar
    If Parent Is Nothing Then
        VScroll.Create Layout, "VScrollBar", L, T, W, H
    Else
        VScroll.Create Parent, "VScrollBar", L, T, W, H
    End If
    ViewList.Add VScroll
    Set CreateVScrollBar = VScroll
End Function

Public Function CreateHScrollBar(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cHScrollBar
    Dim HScroll As New cHScrollBar
    If Parent Is Nothing Then
        HScroll.Create Layout, "HScrollBar", L, T, W, H
    Else
        HScroll.Create Parent, "HScrollBar", L, T, W, H
    End If
    ViewList.Add HScroll
    Set CreateHScrollBar = HScroll
End Function

Public Function CreateSeekBar(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cSeekBar
    Dim SeekBar As New cSeekBar
    If Parent Is Nothing Then
        SeekBar.Create Layout, "SeekBar", L, T, W, H
    Else
        SeekBar.Create Parent, "SeekBar", L, T, W, H
    End If
    ViewList.Add SeekBar
    Set CreateSeekBar = SeekBar
End Function

'Public Function CreateTextBox(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cTextBox
'    Dim TextBox As New cTextBox
'    If Parent Is Nothing Then
'        TextBox.Create Layout, "TextBox", L, T, W, H
'    Else
'        TextBox.Create Parent, "TextBox", L, T, W, H
'    End If
'    ViewList.Add TextBox
'    Set CreateTextBox = TextBox
'End Function

Public Function CreateEditBox(Parent As Object, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As cEditBox
    Dim EditBox As New cEditBox
    If Parent Is Nothing Then
        EditBox.Create Layout, "EditBox", L, T, W, H
    Else
        EditBox.Create Parent, "EditBox", L, T, W, H
    End If
    ViewList.Add EditBox
    Set CreateEditBox = EditBox
End Function
