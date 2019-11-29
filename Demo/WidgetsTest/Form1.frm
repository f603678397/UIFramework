VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4500
   ScaleWidth      =   6000
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Activity                As cActivity
Dim WithEvents Layout       As cLayout
Attribute Layout.VB_VarHelpID = -1
Dim IM                      As cImageManager

Dim WithEvents btnOk        As cButton
Attribute btnOk.VB_VarHelpID = -1
Dim WithEvents btnCancel    As cButton
Attribute btnCancel.VB_VarHelpID = -1
Dim Label1                  As cLabel
Dim WithEvents CheckBox     As cCheckBox
Attribute CheckBox.VB_VarHelpID = -1
Dim Frame1                  As cFrame
Dim WithEvents Option1      As cOption
Attribute Option1.VB_VarHelpID = -1
Dim WithEvents Option2      As cOption
Attribute Option2.VB_VarHelpID = -1
Dim Waiting                 As cWaiting
Dim WithEvents Timer1       As cTimer
Attribute Timer1.VB_VarHelpID = -1
Dim Progress1               As cProgressBar
Dim ImageView               As cImageView
Dim VScroll                 As cVScrollBar
Dim HScroll                 As cHScrollBar
Dim SeekBar                 As cSeekBar
Dim WithEvents EditBox                 As cEditBox
Attribute EditBox.VB_VarHelpID = -1

Private Sub EditBox_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Debug.Print KeyCode, Shift
End Sub

Private Sub Form_Load()
    cCore.Initialize

    Set Activity = cCore.CreateActivity(Me.hWnd)
    Set Layout = Activity.CreateLayout
    Set IM = cCore.GetImageManager
    
    IM.LoadAllImage App.Path & "\res\"
    
    cWidgetManager.BindLayout Layout

    Set btnOk = cWidgetManager.CreateButton(Layout, 260, 260, 60, 30)
    Set btnCancel = cWidgetManager.CreateButton(Layout, 330, 260, 60, 30)
    
    Set Label1 = cWidgetManager.CreateLabel(Layout, 10, 10, 100, 20)
    Set CheckBox = cWidgetManager.CreateCheckBox(Layout, 10, 30, 100, 20)
    Set Frame1 = cWidgetManager.CreateFrame(Layout, 10, 50, 100, 80)
    Set Option1 = cWidgetManager.CreateOption(Frame1, 5, 20, 100, 20)
    Set Option2 = cWidgetManager.CreateOption(Frame1, 5, 50, 100, 20)
    
    Set ImageView = cWidgetManager.CreateImageView(Layout, 10, 140, 20, 20)
    Set SeekBar = cWidgetManager.CreateSeekBar(Layout, 35, 140, 80, 20)
    Set EditBox = cWidgetManager.CreateEditBox(Layout, 10, 170, 100, 50)
    
    Set Waiting = cWidgetManager.CreateWaiting(Layout, 10, 265, 20, 20)
    Set Progress1 = cWidgetManager.CreateProgressBar(Layout, 35, 275, 80, 3)
    
    Set VScroll = cWidgetManager.CreateVScrollBar(Layout, 372, 10, 18, 220)
    Set HScroll = cWidgetManager.CreateHScrollBar(Layout, 10, 230, 360, 18)

    Set Timer1 = New cTimer
    Timer1.Create Me.hWnd
    
    With Label1
        .Caption = "¿Ø¼þÊ¾Àý"
        .FontName = "Î¢ÈíÑÅºÚ"
        .LineAlignCenter = True
        .IsAccent = True
    End With
    
    With CheckBox
        .Caption = "Ö÷Ìâ¿ÉÑ¡"
        .Value = True
        .FontName = "Î¢ÈíÑÅºÚ"
    End With
    
    With Frame1
        .Caption = "Ö÷Ìâ"
        .FontName = "Î¢ÈíÑÅºÚ"
    End With
    
    With Option1
        .Caption = "°µÉ«"
        .FontName = "Î¢ÈíÑÅºÚ"
        .Value = True
    End With
    
    With Option2
        .Caption = "Ç³É«"
        .FontName = "Î¢ÈíÑÅºÚ"
    End With
    
    With Progress1
        .Value = 30
        .SecondValue = 0
    End With
    
    ImageView.Src = "sun_light"
    SeekBar.Value = 50
    
    With EditBox
        .AutoWrap = False
        .FontName = "Î¢ÈíÑÅºÚ"
    End With
    
    VScroll.Max = 50
    HScroll.Max = 50
    
    With btnOk
        .Caption = "È·¶¨"
        .FontName = "Î¢ÈíÑÅºÚ"
        .IsAccent = True
    End With
    
    With btnCancel
        .Caption = "È¡Ïû"
        .FontName = "Î¢ÈíÑÅºÚ"
    End With
    
    Activity.SetLayout Layout
    
    cToast.SetShadown(True).SetFontName("Î¢ÈíÑÅºÚ").SetDuration (1000)
    
    Timer1.Interval = 33
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Release
    cCore.Terminate
End Sub

Private Sub btnOk_Click()
'
EditBox.SelText = "1234"
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub CheckBox_ValueChanged()
    Frame1.Enabled = CheckBox.Value
    EditBox.Enabled = CheckBox.Value
End Sub

Private Sub Option1_ValueChanged(ByVal ByUser As Boolean)
    If Not Option1.Value Or Not ByUser Then Exit Sub
    ImageView.Src = "sun_light"
    cWidgetManager.SetPresetTheme DrakTheme
    cToast.SetTheme DrakTheme
    cToast.MakeText Layout, "ÒÑÇÐ»»Îª°µÉ«Ö÷Ìâ", TF_POS_TOP Or TF_WIDTH_MIN
    cToast.Show
End Sub

Private Sub Option2_ValueChanged(ByVal ByUser As Boolean)
    If Not Option2.Value Or Not ByUser Then Exit Sub
    ImageView.Src = "sun_dark"
    cWidgetManager.SetPresetTheme LightTheme
    cToast.SetTheme LightTheme
    cToast.MakeText Layout, "ÒÑÇÐ»»ÎªÇ³É«Ö÷Ìâ", TF_POS_TOP Or TF_WIDTH_MIN
    cToast.Show
End Sub

Private Sub Timer1_onTime()
    Progress1.Value = Progress1.Value + 1
    If Progress1.Value > Progress1.Max Then Progress1.Value = 0
End Sub
