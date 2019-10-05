VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Drawable Builder"
   ClientHeight    =   9060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   604
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   940
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Activity                As cActivity
Dim WithEvents Layout       As cLayout
Attribute Layout.VB_VarHelpID = -1
Dim WithEvents TitleBar     As cView
Attribute TitleBar.VB_VarHelpID = -1
Dim WithEvents btnClose     As cView
Attribute btnClose.VB_VarHelpID = -1
Dim WithEvents btnMax       As cView
Attribute btnMax.VB_VarHelpID = -1
Dim WithEvents btnMin       As cView
Attribute btnMin.VB_VarHelpID = -1
Dim WorkArea                As cView

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnClose_MouseHover()
    btnClose.BackgroundColor = cColor.FromARGB(255, 230, 20, 35)
End Sub

Private Sub btnClose_MouseLeave()
    btnClose.BackgroundColor = cColor.Transparent
End Sub

Private Sub btnClose_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen As New cPen
    Dim cX As Single, cY As Single
    
    cX = btnClose.Width / 2
    cY = btnClose.Height / 2
    
    Pen.Color = cColor.FromARGB(255, 250, 250, 250)
    
    Canvas.SetSmoothingMode SmoothingModeAntiAlias
    Canvas.DrawLine Pen, cX - 5, cY - 5, cX + 4, cY + 4
    Canvas.DrawLine Pen, cX + 4, cY - 5, cX - 5, cY + 4
End Sub

Private Sub btnMax_Click()
    If IsZoomed(Me.hWnd) Then
        ShowWindow Me.hWnd, SW_RESTORE
    Else
        ShowWindow Me.hWnd, SW_SHOWMAXIMIZED
    End If
End Sub

Private Sub btnMax_MouseHover()
    btnMax.BackgroundColor = cColor.DimGray
End Sub

Private Sub btnMax_MouseLeave()
    btnMax.BackgroundColor = cColor.Transparent
End Sub

Private Sub btnMax_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen As New cPen
    Dim cX As Single, cY As Single
    
    cX = btnClose.Width / 2
    cY = btnClose.Height / 2
    
    Pen.Color = cColor.FromARGB(255, 225, 225, 225)
    
    Canvas.DrawRectangle Pen, cX - 5, cY - 5, 10, 10
    Canvas.DrawLine Pen, cX - 5, cY - 4, cX + 5, cY - 4
End Sub

Private Sub btnMin_Click()
    ShowWindow Me.hWnd, SW_SHOWMINIMIZED
End Sub

Private Sub btnMin_MouseHover()
    btnMin.BackgroundColor = cColor.DimGray
End Sub

Private Sub btnMin_MouseLeave()
    btnMin.BackgroundColor = cColor.Transparent
End Sub

Private Sub btnMin_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen As New cPen
    Dim cX As Single, cY As Single
    
    cX = btnClose.Width / 2
    cY = btnClose.Height / 2
    
    Pen.Color = cColor.FromARGB(255, 225, 225, 225)
    
    Canvas.DrawLine Pen, cX - 5, cY + 3, cX + 5, cY + 3
End Sub

Private Sub Form_Load()
    Set Activity = cCore.CreateActivity(Me.hWnd)
    Set Layout = Activity.CreateLayout
    Layout.BackgroundColor = cColor.FromARGB(255, 65, 65, 65)
    
    Set WorkArea = Layout.CreateView(3, 25, Me.ScaleWidth - 6, Me.ScaleHeight - 28)
    WorkArea.BackgroundColor = cColor.FromARGB(255, 30, 30, 30)
    
    Set TitleBar = Layout.CreateView(3, 3, Me.ScaleWidth - 6, 22)
    Set btnMin = TitleBar.CreateView(TitleBar.Width - 75, 0, 25, 20)
    Set btnMax = TitleBar.CreateView(TitleBar.Width - 50, 0, 25, 20)
    Set btnClose = TitleBar.CreateView(TitleBar.Width - 25, 0, 25, 20)
    
    Activity.SetLayout Layout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set TitleBar = Nothing
    Set Layout = Nothing
    Set Activity = Nothing
    ExitApp
End Sub

Private Sub Layout_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim wParam As Long
    ReleaseCapture
    If X < 3 And Y < 3 Then
        wParam = HTTOPLEFT
    ElseIf X >= 3 And X <= Me.ScaleWidth - 3 And Y < 3 Then
        wParam = HTTOP
    ElseIf X > Me.ScaleWidth - 3 And Y < 3 Then
        wParam = HTTOPRIGHT
    ElseIf X > Me.ScaleWidth - 3 And Y >= 3 And Y <= Me.ScaleHeight - 3 Then
        wParam = HTRIGHT
    ElseIf X > Me.ScaleWidth - 3 And Y > Me.ScaleHeight - 3 Then
        wParam = HTBOTTOMRIGHT
    ElseIf X >= 3 And X <= Me.ScaleWidth - 3 And Y > Me.ScaleHeight - 3 Then
        wParam = HTBOTTOM
    ElseIf X < 3 And Y > Me.ScaleHeight - 3 Then
        wParam = HTBOTTOMLEFT
    ElseIf X < 3 And Y >= 3 And Y <= Me.ScaleHeight - 3 Then
        wParam = HTLEFT
    End If
    SendMessageA Me.hWnd, WM_NCLBUTTONDOWN, wParam, 0
End Sub

Private Sub Layout_MouseLeave()
    SetCursor Cursor.Normal
End Sub

Private Sub Layout_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    If (X < 3 And Y < 3) Or (X > Me.ScaleWidth - 3 And Y > Me.ScaleHeight - 3) Then
        SetCursor Cursor.NW_SE
    End If
    
    If (X < 3 And Y > Me.ScaleHeight - 3) Or (X > Me.ScaleWidth - 3 And Y < 3) Then
        SetCursor Cursor.NE_SW
    End If
    
    If X >= 3 And X <= Me.ScaleWidth - 3 Then SetCursor Cursor.NS
    
    If Y >= 3 And Y <= Me.ScaleHeight - 3 Then SetCursor Cursor.WE
End Sub

Private Sub Layout_Resize()
    If IsIconic(Me.hWnd) Then Exit Sub
    
    If Me.ScaleWidth < 200 Then Me.Width = 200 * Screen.TwipsPerPixelX
    If Me.ScaleHeight < 150 Then Me.Height = 150 * Screen.TwipsPerPixelY
    
    TitleBar.Move 3, 3, Me.ScaleWidth - 6, 22
    btnMin.Move TitleBar.Width - 75, 0, 25, 20
    btnMax.Move TitleBar.Width - 50, 0, 25, 20
    btnClose.Move TitleBar.Width - 25, 0, 25, 20
    WorkArea.Move 3, 25, Me.ScaleWidth - 6, Me.ScaleHeight - 28
End Sub

Private Sub TitleBar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    ReleaseCapture
    SendMessageA Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub TitleBar_Paint(Canvas As Drawing2D.cGraphics)
    Dim Family          As New cFontFamily
    Dim Font            As New cFont
    Dim strFormat       As New cStringFormat
    Dim Brush           As New cSolidBrush
    Dim Pen             As New cPen
    
    Family.FromName "Î¢ÈíÑÅºÚ"
    Font.FromFamily Family, 12, FontStyleRegular, UnitPixel
    strFormat.Align = StringAlignmentNear
    strFormat.LineAlign = StringAlignmentCenter
    Pen.Color = cColor.FromARGB(255, 215, 215, 215)
    Brush.Color = cColor.FromARGB(255, 215, 215, 215)
    
    Canvas.SetTextRenderingHint TextRenderingHintClearTypeGridFit
    Canvas.Clear
    Canvas.DrawRectangle Pen, 3, 3, 8, 8
    Canvas.FillRectangle Brush.GetBaseBrush, 7, 7, 8, 8
    Canvas.DrawString Me.Caption, Font, NewRectF(18, 0, TitleBar.Width - 20, TitleBar.Height - 2), strFormat, Brush.GetBaseBrush
End Sub
