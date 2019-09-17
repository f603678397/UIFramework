VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToast"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents View         As cView
Attribute View.VB_VarHelpID = -1
Private WithEvents Timer        As cTimer
Attribute Timer.VB_VarHelpID = -1
Private Context                 As cLayout
Private Msg                     As String
Private mFlag                   As ToastFlag
Private mBKColor                As Long
Private mFGColor                As Long
Private mFontName               As String
Private mFontSize               As Single
Private bShadown                As Boolean

Public Enum ToastFlag
    TF_POS_TOP = 0
    TF_POS_BOTTOM = 1
    TF_WIDTH_MAX = 0
    TF_WIDTH_MIN = 2
End Enum

Private Const TipHeight As Integer = 25

Public Function SetFontName(ByVal FontName As String) As cToast
    mFontName = Trim(FontName)
    Set SetFontName = Me
End Function

Public Function SetFontSize(ByVal fSize As Single) As cToast
    mFontSize = fSize
    Set SetFontSize = Me
End Function

Public Function SetShadown(ByVal bOpen As Boolean) As cToast
    bShadown = bOpen
    Set SetShadown = Me
End Function

Public Function SetBackgroundColor(ByVal nColor As Long) As cToast
    mBKColor = nColor
    Set SetBackgroundColor = Me
End Function

Public Function SetForegroundColor(ByVal nColor As Long) As cToast
    mFGColor = nColor
    Set SetForegroundColor = Me
End Function

Public Function MakeText(Layout As cLayout, ByVal Message As String, Flag As ToastFlag) As cToast
    If Layout Is Nothing Then Exit Function
    
    Set Context = Layout
    Msg = Message
    mFlag = Flag
    Set MakeText = Me
End Function

Public Sub Show(Optional ByVal Offset As Integer = 0)
    Dim Activity As cActivity
    Dim Left As Integer, Top As Integer, Width As Integer, Height As Integer
    
    If Context Is Nothing Then Exit Sub
    Set Activity = Context.GetRoot
    If Activity Is Nothing Then Exit Sub
    
    Top = Offset
    Width = Activity.Width
    Height = Activity.Height
    
    If (mFlag And TF_WIDTH_MIN) = TF_WIDTH_MIN Then
        If Width > 180 Then
            Left = (Width - 180) / 2
            Width = 180
        End If
    End If
    
    If (mFlag And TF_POS_BOTTOM) = TF_POS_BOTTOM Then
        If Height > TipHeight Then
            Top = Height - TipHeight - Offset
        End If
    End If
    Height = TipHeight
    
    If mBKColor = 0 Then mBKColor = cColor.DarkGray
    If mFGColor = 0 Then mFGColor = cColor.White
    If mFontName = vbNullString Then mFontName = "ËÎÌו"
    If mFontSize <= 0 Then mFontSize = 15
    
    Set View = Context.CreateView(Left, Top, Width, Height)

    Set Timer = New cTimer
    Timer.Create Activity.hWnd
    Timer.Interval = 33
End Sub

Private Sub Timer_onTime()
    Static Opacity As Single
    Static B As Byte
    Static T1 As Long
    Dim T2 As Long
    Dim fStep As Single
    
    fStep = 0.2
    
    Select Case B
        Case 0
            Opacity = Opacity + fStep
            If Opacity >= 1 Then
                Opacity = 1
                B = 1
                T1 = GetTickCount
            End If
        Case 1
            T2 = GetTickCount
            If T2 - T1 >= 1200 Then '
                B = 2
            End If
        Case 2
            Opacity = Opacity - fStep
            If Opacity <= 0 Then
                B = 0
                Timer.Enabled = False
                Timer.Release
                Context.RemoveView View
                
                Set View = Nothing
                Set Context = Nothing
                Set Timer = Nothing
                Exit Sub
            End If
    End Select
    
    View.Opacity = Opacity
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Brush As New cSolidBrush
    Dim Font As New cFont
    Dim Format As New cStringFormat
    Dim Family As New cFontFamily
    Dim RECT As Drawing2D.RECTF
    Dim Pen As New cPen
    Dim Pos As Byte
    Dim i As Integer
    
    Brush.Color = mFGColor
    
    Family.FromName mFontName
    Font.FromFamily Family, mFontSize, FontStyleBold, UnitPixel
    
    If (mFlag And TF_WIDTH_MIN) = TF_WIDTH_MIN Then
        Format.Align = StringAlignmentCenter
    Else
        Format.Align = StringAlignmentNear
    End If
    Format.LineAlign = StringAlignmentCenter
    Format.Trimming = StringTrimmingEllipsisCharacter
    Format.Flags = StringFormatFlagsNoWrap
    
    If bShadown Then
        If (mFlag And TF_POS_TOP) = TF_POS_TOP Then
            If (mFlag And TF_WIDTH_MAX) = TF_WIDTH_MAX Then RECT = NewRectF(0, 0, View.Width, View.Height - 3): Pos = 0
            If (mFlag And TF_WIDTH_MIN) = TF_WIDTH_MIN Then RECT = NewRectF(3, 0, View.Width - 6, View.Height - 3): Pos = 1
        End If
        
        If (mFlag And TF_POS_BOTTOM) = TF_POS_BOTTOM Then
            If (mFlag And TF_WIDTH_MAX) = TF_WIDTH_MAX Then RECT = NewRectF(0, 3, View.Width, View.Height - 3): Pos = 2
            If (mFlag And TF_WIDTH_MIN) = TF_WIDTH_MIN Then RECT = NewRectF(3, 3, View.Width - 6, View.Height - 3): Pos = 3
        End If

        For i = 3 To 0 Step -1
            Pen.Color = cColor.FromARGB(i ^ 2 * 20, 0, 0, 0)
            
            Select Case Pos
                Case 0
                    Canvas.DrawLine Pen, 0, View.Height - i, View.Width, View.Height - i
                Case 1
                    Canvas.DrawLine Pen, i, View.Height - i, View.Width - i, View.Height - i
                    Canvas.DrawLine Pen, i, 0, i, View.Height - i - 1
                    Canvas.DrawLine Pen, View.Width - i, 0, View.Width - i, View.Height - i - 1
                Case 2
                    Canvas.DrawLine Pen, 0, i, View.Width, i
                Case 3
                    Canvas.DrawLine Pen, i, i, View.Width - i, i
                    Canvas.DrawLine Pen, i, i, i, View.Height - 1
                    Canvas.DrawLine Pen, View.Width - i, i, View.Width - i, View.Height - 1
            End Select
        Next
    Else
        RECT = NewRectF(0, 0, View.Width, View.Height)
    End If
    
    Canvas.SetTextRenderingHint TextRenderingHintClearTypeGridFit
    Canvas.SetClip RECT.Left, RECT.Top, RECT.Right, RECT.bottom, CombineModeReplace
    Canvas.Clear mBKColor
    Canvas.DrawString Msg, Font, RECT, Format, Brush.GetBaseBrush
    Canvas.ResetClip
End Sub
