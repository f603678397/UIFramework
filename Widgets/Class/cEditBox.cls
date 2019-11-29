VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEditBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents View         As cView
Attribute View.VB_VarHelpID = -1
Private Bitmap                  As New cBitmap

Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseHover()
Public Event MouseLeave()
Public Event GotFocus()
Public Event LostFocus()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event KeyPress(ByVal KeyAscii As Long)
Public Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event TextChanged()

Private mText       As String
Private OutText     As String
Private mFontName   As String
Private mFontSize   As Single
Private mStyle      As FontStyle
Private mAccent     As Boolean
Private mAutoWrap   As Boolean
Private mSingle     As Boolean

Private bFocus      As Boolean
Private bDown       As Boolean
Private bChanged    As Boolean
Private bWheel      As Boolean

Private ThemeColor  As ThemeColorType

Private CharRange   As New cArrayList
Private InputPos    As Drawing2D.RECTF

Private CharIndex   As Long
Private CharLast    As Long
Private DownIndex   As Long

Private MI_IBEAM    As Long
Private MI_NORMAL   As Long

Private OffsetX     As Long
Private OffsetY     As Long

Private Sub Class_Initialize()
    mText = "TextBox"
    mFontName = "宋体"
    mFontSize = 14
    mStyle = FontStyleRegular
    bChanged = True
    mAutoWrap = True
    
    CharLast = -1
    
    MI_IBEAM = LoadCursorWL(ByVal 0&, IDC_IBEAM)
    MI_NORMAL = LoadCursorWL(ByVal 0&, IDC_ARROW)
    
    SetTheme Preset
    
    ThemeColor.BKColor = cColor.White
End Sub

Public Sub Create(Parent As Object, ByVal sCaption As String, ByVal nLeft As Integer, ByVal nTop As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer)
    Set View = Parent.CreateView(nLeft, nTop, nWidth, nHeight)
    mText = sCaption
    View.Name = "TextBox"
End Sub

Public Sub BindView(DstView As cView)
    Dim V As cView
    
    If DstView Is View Then Exit Sub
    
    Set V = View
    Set View = Nothing
    Set View = DstView
    
    bFocus = View.IsFocus
    bDown = False
    mText = "TextBox"
    
    V.Refresh
    View.Refresh
End Sub

Public Function GetView() As cView
    Set GetView = View
End Function
'=========================================================================================
'隐藏函数
Public Sub SetAttr(ByVal AttrName As String, ByVal Value As String)
    Select Case LCase(AttrName)
    '-----------------------------------------
        Case LCase("Name")
            Me.Name = Value
        Case LCase("Left")
            Me.left = CInt(Value)
        Case LCase("Top")
            Me.Top = CInt(Value)
        Case LCase("Width")
            Me.Width = CInt(Value)
        Case LCase("Height")
            Me.Height = CInt(Value)
        Case LCase("Enabled")
            Me.Enabled = CBool(Value)
        Case LCase("Visible")
            Me.Visible = CBool(Value)
        Case LCase("Opacity")
            Me.Opacity = CSng(Value)
    '-----------------------------------------
        Case LCase("Text")
            Me.Text = Value
        Case LCase("BackColor")
            Me.BackColor = CLng(Value)
        Case LCase("ForeColor")
            Me.ForeColor = CLng(Value)
        Case LCase("IsAccent")
            Me.IsAccent = CBool(Value)
        Case LCase("AccentColor")
            Me.AccentColor = CLng(Value)
        Case LCase("BorderColor")
            Me.BorderColor = CLng(Value)
    '-----------------------------------------
        Case LCase("FontBold")
            Me.FontBold = CBool(Value)
        Case LCase("FontItalic")
            Me.FontItalic = CBool(Value)
        Case LCase("FontName")
            Me.FontName = Value
        Case LCase("FontSize")
            Me.FontSize = CSng(Value)
        Case LCase("FontStrikeout")
            Me.FontStrikeout = CBool(Value)
        Case LCase("FontUnderLine")
            Me.FontUnderLine = CBool(Value)
    End Select
End Sub

'=========================================================================================
'通用代码

Public Property Get Name() As String: Name = View.Name: End Property
Public Property Let Name(ByVal sName As String): View.Name = Trim(sName): End Property

Public Property Get left() As Integer: left = View.left: End Property
Public Property Let left(ByVal nLeft As Integer)
    If View.left = nLeft Then Exit Property
    View.left = nLeft
End Property

Public Property Get Top() As Integer: Top = View.Top: End Property
Public Property Let Top(ByVal nTop As Integer)
    If View.Top = nTop Then Exit Property
    View.Top = nTop
End Property

Public Property Get Width() As Integer: Width = View.Width: End Property
Public Property Let Width(ByVal nWidth As Integer)
    If View.Width = nWidth Then Exit Property
    View.Width = nWidth
End Property

Public Property Get Height() As Integer: Height = View.Height: End Property
Public Property Let Height(ByVal nHeight As Integer)
    If View.Height = nHeight Then Exit Property
    View.Height = nHeight
End Property

Public Property Get Tag() As String: Tag = View.Tag: End Property
Public Property Let Tag(ByVal sTag As String): View.Tag = sTag: End Property

Public Property Get Enabled() As Boolean: Enabled = View.Enabled: End Property
Public Property Let Enabled(ByVal bEnabled As Boolean)
    If View.Enabled = bEnabled Then Exit Property
    View.Enabled = bEnabled
    bChanged = True
    View.Refresh
End Property

Public Property Get Visible() As Boolean: Visible = View.Visible: End Property
Public Property Let Visible(ByVal bVisible As Boolean)
    If View.Visible = bVisible Then Exit Property
    View.Visible = bVisible
End Property

Public Property Get Opacity() As Single: Opacity = View.Opacity: End Property
Public Property Let Opacity(ByVal fOpacity As Single)
    If View.Opacity = fOpacity Then Exit Property
    View.Opacity = fOpacity
End Property

Public Property Get ZOrder() As Long: ZOrder = View.ZOrder: End Property
Public Property Let ZOrder(ByVal nZ As Long)
    If View.ZOrder = nZ Then Exit Property
    View.ZOrder = nZ
End Property

Public Sub Move(ByVal nLeft As Integer, Optional nTop, Optional nWidth, Optional nHeight)
    Dim T As Integer, W As Integer, H As Integer
    
    T = IIf(IsMissing(nTop), View.Top, CInt(nTop))
    W = IIf(IsMissing(nWidth), View.Width, CInt(nWidth))
    H = IIf(IsMissing(nHeight), View.Height, CInt(nHeight))
    
    View.Move nLeft, T, W, H
End Sub

Public Property Get BackColor() As Long: BackColor = ThemeColor.BKColor: End Property
Public Property Let BackColor(ByVal nColor As Long)
    If nColor = ThemeColor.BKColor Then Exit Property
    
    ThemeColor.BKColor = nColor
    ThemeColor.BKDrakColor = cColor.SetBrightness(nColor, -0.1)
    ThemeColor.BKLightColor = cColor.SetBrightness(nColor, 0.1)
    
    View.Refresh
End Property

Public Property Get ForeColor() As Long: ForeColor = ThemeColor.TextColor: End Property
Public Property Let ForeColor(ByVal nColor As Long)
    If nColor = ThemeColor.TextColor Then Exit Property
    
    ThemeColor.TextColor = nColor
    View.Refresh
End Property

Public Sub Refresh(): View.Refresh: End Sub

Public Sub SetTheme(NewTheme As PresetThemeEnum)
    Select Case NewTheme
        Case PresetThemeEnum.DrakTheme
            SetDarkTheme ThemeColor
        Case PresetThemeEnum.LightTheme
            SetLightTheme ThemeColor
    End Select
    ThemeColor.TextColor = cColor.Black
End Sub

'=========================================================================================

Public Property Get Text() As String: Text = mText: End Property
Public Property Let Text(ByVal sText As String)
    Dim i As Long
    CharLast = -1
    
    If sText = mText Then Exit Property
    mText = sText
    bChanged = True
    View.Refresh
    RaiseEvent TextChanged
End Property

Public Property Get FontName() As String: FontName = mFontName: End Property
Public Property Let FontName(ByVal sFontName As String)
    If mFontName = sFontName Then Exit Property
    mFontName = sFontName
    bChanged = True
    View.Refresh
End Property

Public Property Get FontSize() As Single: FontSize = mFontSize: End Property
Public Property Let FontSize(ByVal nSize As Single)
    If mFontSize = nSize Then Exit Property
    mFontSize = nSize
    bChanged = True
    View.Refresh
End Property

Public Property Get FontBold() As Boolean
    If (mStyle And FontStyleBold) = FontStyle.FontStyleBold Then FontBold = True
End Property
Public Property Let FontBold(ByVal bBold As Boolean)
    Dim nStyle As FontStyle
    
    nStyle = (mStyle Or FontStyleBold)
    If Not bBold Then
        nStyle = nStyle Xor FontStyleBold
    End If
    
    If mStyle = nStyle Then Exit Property
    mStyle = nStyle
    bChanged = True
    View.Refresh
End Property

Public Property Get FontItalic() As Boolean
    If (mStyle And FontStyleItalic) = FontStyle.FontStyleItalic Then FontItalic = True
End Property
Public Property Let FontItalic(ByVal bItalic As Boolean)
    Dim nStyle As FontStyle
    
    nStyle = (mStyle Or FontStyleItalic)
    If Not bItalic Then
        nStyle = nStyle Xor FontStyleItalic
    End If
    
    If mStyle = nStyle Then Exit Property
    mStyle = nStyle
    bChanged = True
    View.Refresh
End Property

Public Property Get FontUnderLine() As Boolean
    If (mStyle And FontStyleUnderline) = FontStyle.FontStyleUnderline Then FontUnderLine = True
End Property
Public Property Let FontUnderLine(ByVal bLine As Boolean)
    Dim nStyle As FontStyle
    
    nStyle = (mStyle Or FontStyleUnderline)
    If Not bLine Then
        nStyle = nStyle Xor FontStyleUnderline
    End If
    
    If mStyle = nStyle Then Exit Property
    mStyle = nStyle
    bChanged = True
    View.Refresh
End Property

Public Property Get FontStrikeout() As Boolean
    If (mStyle And FontStyleStrikeout) = FontStyle.FontStyleStrikeout Then FontStrikeout = True
End Property
Public Property Let FontStrikeout(ByVal bStru As Boolean)
    Dim nStyle As FontStyle
    
    nStyle = (mStyle Or FontStyleStrikeout)
    If Not bStru Then
        nStyle = nStyle Xor FontStyleStrikeout
    End If
    
    If mStyle = nStyle Then Exit Property
    mStyle = nStyle
    bChanged = True
    View.Refresh
End Property

Public Property Get IsAccent() As Boolean: IsAccent = mAccent: End Property
Public Property Let IsAccent(ByVal bAccent As Boolean)
    If mAccent = bAccent Then Exit Property
    mAccent = bAccent
    View.Refresh
End Property

Public Property Get AccentColor() As Long: AccentColor = ThemeColor.AccentColor: End Property
Public Property Let AccentColor(ByVal nColor As Long)
    If ThemeColor.AccentColor = nColor Then Exit Property
    ThemeColor.AccentColor = nColor
    ThemeColor.AccentDarkColor = cColor.SetBrightness(nColor, -0.2)
    ThemeColor.AccentLightColor = cColor.SetBrightness(nColor, 0.2)
    View.Refresh
End Property

Public Property Get BorderColor() As Long: BorderColor = ThemeColor.BorderColor: End Property
Public Property Let BorderColor(ByVal nColor As Long)
    If nColor = ThemeColor.BorderColor Then Exit Property
    
    ThemeColor.BorderColor = nColor
    
    View.Refresh
End Property

Public Property Get AutoWrap() As Boolean: AutoWrap = mAutoWrap: End Property
Public Property Let AutoWrap(ByVal bAuto As Boolean)
    If bAuto = mAutoWrap Then Exit Property
    mAutoWrap = bAuto
    bChanged = True
    View.Refresh
End Property

Public Property Get SingleLine() As Boolean: SingleLine = mSingle: End Property
Public Property Let SingleLine(ByVal bSingle As Boolean)
    If bSingle = mSingle Then Exit Property
    
    mSingle = bSingle
    mAutoWrap = Not mSingle
    
    If mSingle Then mText = Replace(mText, vbCrLf, "")
    
    View.Refresh
End Property

Public Property Get SelStart() As Long: SelStart = CharIndex: End Property
Public Property Let SelStart(ByVal nStart As Long)
    If nStart = CharIndex Then Exit Property
    
    CharIndex = nStart
    View.Refresh
End Property

Public Property Get SelText() As String
    If CharLast > -1 Then
        SelText = mID$(mText, CharIndex + 1, CharLast - CharIndex + 1)
    End If
End Property
Public Property Let SelText(ByVal sText As String)
    Dim s1 As String
    Dim s2 As String
    
    If CharIndex > 0 Then s1 = left$(OutText, CharIndex)
    
    If CharLast > -1 Then
        s2 = Right$(OutText, Len(OutText) - CharLast - 1)
    Else
        s2 = Right$(OutText, Len(OutText) - CharIndex)
    End If
    
    CharIndex = CharIndex + Len(sText)
    Me.Text = s1 & sText & s2
End Property

Public Property Get SelLength() As Long
    If CharLast = -1 Then
        SelLength = 0
    ElseIf CharIndex = 0 Then
        SelLength = Len(OutText)
    Else
        SelLength = CharLast - CharIndex + 1
    End If
    View.Refresh
End Property
Public Property Let SelLength(ByVal nLen As Long)
    If nLen < 0 Then: CharLast = -1: GoTo ReDraw
    If CharIndex = 0 Then
        CharLast = nLen
    Else
        CharLast = CharIndex + nLen - 1
    End If
    
ReDraw:
    View.Refresh
End Property

Private Sub View_Click(): RaiseEvent Click: End Sub
Private Sub View_DblClick(): RaiseEvent DblClick: End Sub

Private Sub View_GotFocus()
    RaiseEvent GotFocus
    bFocus = True
    View.Refresh
End Sub

Private Sub View_LostFocus()
    RaiseEvent LostFocus
    bFocus = False
    View.Refresh
End Sub

Private Sub View_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Dim s As String
    RaiseEvent KeyDown(KeyCode, Shift)
    Select Case KeyCode
        Case vbKeyLeft
            If CharIndex > 0 Then CharIndex = CharIndex - 1
            If CharLast <> -1 Then
                CharLast = -1
                bChanged = True
            End If
            View.Refresh
        Case vbKeyRight
            If CharIndex < CharRange.GetCount - 1 Then CharIndex = CharIndex + 1
            If CharLast <> -1 Then
                CharLast = -1
                bChanged = True
            End If
            View.Refresh
        Case vbKeyA
            If Shift = vbCtrlMask Then
                CharIndex = 0
                CharLast = Len(OutText) - 1
                bChanged = True
                View.Refresh
            End If
        Case vbKeyC
            If Shift = vbCtrlMask Then
                Clipboard.SetText Me.SelText
            End If
        Case vbKeyX
            If Shift = vbCtrlMask Then
                Clipboard.Clear
                Clipboard.SetText Me.SelText
            End If
        Case vbKeyV
            If Shift = vbCtrlMask Then
                s = Clipboard.GetText
                s = Replace(s, vbCrLf, vbLf)
                If s <> vbNullString Then
                    Me.SelText = s
                End If
            End If
    End Select
End Sub

Private Sub View_KeyPress(ByVal KeyAscii As Long)
    Dim sChar   As String
    Dim sText   As String
    Dim s1      As String
    Dim s2      As String
    
    RaiseEvent KeyPress(KeyAscii)
    
    Select Case KeyAscii
        Case vbKeyBack
            If CharLast = -1 Then
                If CharIndex > 0 And CharIndex <= Len(OutText) Then
                    s1 = left$(OutText, CharIndex - 1)
                    s2 = Right$(OutText, Len(OutText) - CharIndex)
                    sText = s1 & s2
                    CharIndex = CharIndex - 1
                ElseIf CharIndex > Len(OutText) Then
                    sText = left$(OutText, Len(OutText) - 1)
                    CharIndex = CharIndex - 1
                End If
            Else
                Me.SelText = ""
                Exit Sub
            End If
        Case vbKeyReturn
            If CharLast = -1 Then
                sChar = vbLf
                If CharIndex <= Len(OutText) Then
                    s1 = left$(OutText, CharIndex)
                    s2 = Right$(OutText, Len(OutText) - CharIndex)
                    sText = s1 & sChar & s2
                    CharIndex = CharIndex + 1
                End If
            Else
                Me.SelText = vbCrLf
                Exit Sub
            End If
        Case Else
            If KeyAscii >= 0 And KeyAscii < 32 Then
                Exit Sub
            Else
                sChar = Chr$(KeyAscii)
                If CharLast = -1 Then
                    If CharIndex <= Len(OutText) Then
                        s1 = left$(OutText, CharIndex)
                        s2 = Right$(OutText, Len(OutText) - CharIndex)
                        sText = s1 & sChar & s2
                        CharIndex = CharIndex + 1
                    End If
                Else
                    Me.SelText = sChar
                    Exit Sub
                End If
            End If
    End Select

    sText = Replace(sText, vbLf, vbCrLf)
    Me.Text = sText
End Sub

Private Sub View_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub View_MouseHover()
    RaiseEvent MouseHover
    SetCursor MI_IBEAM
End Sub
Private Sub View_MouseLeave()
    RaiseEvent MouseLeave
    SetCursor MI_NORMAL
End Sub

Private Sub View_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button <> vbLeftButton Then Exit Sub
    bDown = True
    CharLast = -1
    
    If X >= 2 And X <= View.Width - 4 And Y >= 2 And Y <= View.Height - 4 Then
        CharIndex = GetCharIndex(X, Y)
        DownIndex = CharIndex
        bChanged = True
        View.Refresh
    End If
End Sub

Private Sub View_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim Index As Long
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If bDown Then
        Index = GetCharIndex(X, Y)
        If Index = CharRange.GetCount - 1 Then Index = CharRange.GetCount - 2
        If Index < DownIndex Then
            CharIndex = Index
            CharLast = DownIndex - 1
        Else
            CharIndex = DownIndex
            CharLast = Index
        End If
        
        bChanged = True
        View.Refresh
    End If
End Sub

Private Sub View_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    bDown = False
End Sub

Private Sub View_MouseWheel(ByVal Forward As Boolean)
    If Bitmap.Height < View.Height - 4 Then Exit Sub
    
    If Forward Then
        If OffsetY > 0 Then
            OffsetY = 0
        ElseIf OffsetY < 0 Then
            OffsetY = OffsetY + 5
        Else
            Exit Sub
        End If
    Else
        If (View.Height - 4) - OffsetY > Bitmap.Height Then
            OffsetY = View.Height - 4 - Bitmap.Height
        ElseIf (View.Height - 4) - OffsetY < Bitmap.Height Then
            OffsetY = OffsetY - 5
        Else
            Exit Sub
        End If
    End If
    bWheel = True
    View.Refresh
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen         As New cPen
    Dim txtPen      As New cPen
    Dim Brush       As New cSolidBrush
    Dim BKColor     As Long

    If View.Enabled Then
        If mAccent Or bFocus Then
            Pen.Color = ThemeColor.AccentColor
        Else
            Pen.Color = ThemeColor.BorderColor
        End If
        Brush.Color = ThemeColor.TextColor
        txtPen.Color = cColor.Black
        BKColor = ThemeColor.BKColor
    Else
        Pen.Color = ThemeColor.BorderDisEnableColor
        Brush.Color = ThemeColor.BorderColor
        txtPen.Color = ThemeColor.BorderDisEnableColor
        BKColor = ThemeColor.BorderDisEnableColor
    End If
    
    If bChanged Then
        DrawText Brush.GetBaseBrush, BKColor
    End If
    
    If Not bWheel Then GetInputPosition
    
    Canvas.Clear BKColor
    Canvas.SetClip 2, 2, View.Width - 4, View.Width - 4, CombineModeReplace
    Canvas.DrawImage Bitmap.ToImage, 2 + OffsetX, 2 + OffsetY
    
    If bFocus And View.Enabled Then
        With InputPos
            Canvas.DrawLine txtPen, .left + OffsetX, .Top + OffsetY, .left + OffsetX, .Top + .Bottom + OffsetY
        End With
    End If
    
    Canvas.ResetClip
    
    Canvas.DrawRectangle Pen, 0, 0, View.Width - 1, View.Height - 1
    bWheel = False
End Sub

Private Sub DrawText(Brush As cBrush, ByVal BKColor As Long)
    Dim Fam             As New cFontFamily
    Dim Font            As New cFont
    Dim strFormat       As New cStringFormat
    Dim selBrush        As New cSolidBrush
    Dim SelTextBrush    As New cSolidBrush
    Dim ts(0 To 3)      As Single
    
    Dim Canvas          As cGraphics
    Dim BmpSize         As Drawing2D.RECTF
    
    Dim sText           As String
    Dim i               As Long
    
    Dim nFirst          As Long
    Dim nLen            As Long
    Dim txtLen          As Long
    Dim sChar           As String
    Dim CharRect        As Drawing2D.RECTF
    
    sText = Replace(mText, vbCrLf, vbLf)
    sText = Replace(sText, vbCr, vbLf)
    
    For i = 0 To UBound(ts)
        ts(i) = 50
    Next
    
    Fam.FromName mFontName
    Font.FromFamily Fam, mFontSize, mStyle, UnitPixel
    strFormat.Flags = StringFormatFlagsNoClip Or StringFormatFlagsMeasureTrailingSpaces
    If Not mAutoWrap Then strFormat.Flags = strFormat.Flags Or StringFormatFlagsNoWrap
    strFormat.SetTabStops 0, ts
    
    Bitmap.CreateEmptyBitmap View.Width - 4, MAX_INT
    BmpSize = GetBitmapSize(sText, Bitmap.GetGraphics, Font, strFormat)
    Bitmap.CreateEmptyBitmap BmpSize.Right, BmpSize.Bottom
    Set Canvas = Bitmap.GetGraphics
    
    txtLen = Len(sText)
    CharRange.Clear
    If txtLen = 0 Then
        CharRange.Add NewRectF(2, 2, 1, mFontSize)
    Else
        If txtLen < 33 Then
            nLen = txtLen
            GetCharRange sText, nFirst, nLen, Canvas, Font, strFormat, BmpSize
        Else
            Do
                If txtLen - nFirst > 32 Then
                    nLen = 32
                Else
                    nLen = txtLen - nFirst
                End If
                
                GetCharRange sText, nFirst, nLen, Canvas, Font, strFormat, BmpSize
                nFirst = nFirst + nLen
            Loop While nFirst < txtLen
        End If
        
        sChar = Right$(sText, 1)
        CharRect = CharRange.GetValue(CharRange.GetCount - 1)
        
        If sChar = vbLf Then
            CharRect.Top = CharRect.Top + CharRect.Bottom
            CharRect.left = 2
        Else
            CharRect.left = CharRect.left + CharRect.Right
        End If
        CharRect.Right = 1
        CharRange.Add CharRect
    End If
    
    Canvas.Clear BKColor
    
    Canvas.SetTextRenderingHint TextRenderingHintClearTypeGridFit
    Canvas.DrawString sText, Font, BmpSize, strFormat, Brush
    
    If CharLast > -1 And View.Enabled Then
        selBrush.Color = ThemeColor.AccentColor
        SelTextBrush.Color = cColor.White
        For i = CharIndex To CharLast
            CharRect = CharRange.GetValue(i)
            CharRect.left = CharRect.left - 2
            CharRect.Top = CharRect.Top - 2
            
            Canvas.FillRectangle selBrush.GetBaseBrush, CharRect.left, CharRect.Top, CharRect.Right, CharRect.Bottom
        Next
        
        For i = CharIndex To CharLast
            CharRect = CharRange.GetValue(i)
            CharRect.left = CharRect.left - 2
            CharRect.Top = CharRect.Top - 2
            
            If i < CharRange.GetCount - 1 Then
                Canvas.DrawString mID$(sText, i + 1, 1), Font, CharRect, strFormat, SelTextBrush.GetBaseBrush
            End If
        Next
    End If
    
    OutText = sText
    bChanged = False
End Sub

Private Function GetBitmapSize(ByVal sText As String, Canvas As Drawing2D.cGraphics, Font As cFont, strFormat As cStringFormat) As Drawing2D.RECTF
    Dim Bound       As Drawing2D.RECTF
    
    If mAutoWrap Then Bound.Right = View.Width - 4
    
    Bound = Canvas.MeasureString(sText, Font, Bound, strFormat)
    Bound.Bottom = Bound.Bottom + 2
    GetBitmapSize = Bound
End Function

Private Function GetCharIndex(ByVal X As Long, ByVal Y As Long) As Long
    Dim i           As Long
    Dim CharRect    As Drawing2D.RECTF
    Dim bFindLine   As Boolean
    Dim Index       As Long
    
    X = X - OffsetX
    Y = Y - OffsetY
    
    For i = 0 To CharRange.GetCount - 2
        CharRect = CharRange.GetValue(i)
        With CharRect
            If Y >= .Top And Y <= .Top + .Bottom Then
                Index = i
                bFindLine = True
                If X >= .left And X <= .left + .Right Then
                    If X - .left < .Right / 2 Then
                        Index = i
                    Else
                        If mID$(OutText, i + 1, 1) = vbLf Then
                            Index = i
                        Else
                            Index = i + 1
                        End If
                    End If
                    GoTo ReturnValue
                End If
            End If
            
            If Not bFindLine And Y < .Top Then
                Index = i - 1
                bFindLine = True
            End If
        End With
    Next
    
    If Not bFindLine Or Index = CharRange.GetCount - 2 Then
        Index = CharRange.GetCount - 1
    End If
    
ReturnValue:
    GetCharIndex = Index
End Function

Private Function GetInputPosition()
    If CharLast = -1 Then
        InputPos = CharRange.GetValue(CharIndex)
    Else
        If CharIndex < DownIndex Then
            InputPos = CharRange.GetValue(CharIndex)
        Else
            InputPos = CharRange.GetValue(CharLast + 1)
        End If
    End If
    
    If InputPos.left + OffsetX < 2 Or InputPos.left + InputPos.Right + OffsetX > View.Width - 4 Then
        If InputPos.left + InputPos.Right < View.Width - 4 Then
            OffsetX = 0
        Else
            OffsetX = (View.Width - 4) - (InputPos.left + InputPos.Right)
        End If
    End If
    
    If InputPos.Top + OffsetY < 2 Or InputPos.Top + InputPos.Bottom + OffsetY > View.Height - 4 Then
        If InputPos.Top + InputPos.Bottom <= View.Height - 4 Then
            OffsetY = 0
        Else
            OffsetY = (View.Height - 4) - (InputPos.Top + InputPos.Bottom)
        End If
    End If
    
    SetCompositionWindowPosition InputPos.left + OffsetX, InputPos.Top + OffsetY
End Function

Private Sub SetCompositionWindowPosition(ByVal X As Integer, ByVal Y As Integer)
    Dim HIMC        As Long
    Dim Comp        As COMPOSITIONFORM
    Dim hWnd        As Long
    
    hWnd = View.GetLayout.GetActivity.hWnd
    HIMC = ImmGetContext(hWnd)
    With Comp
        .dwStyle = CFS_POINT Or CFS_RECT
        .ptCurrentPos.X = View.AbsLeft + X
        .ptCurrentPos.Y = View.AbsTop + Y
        .rcArea.left = View.AbsLeft + 2
        .rcArea.Top = View.AbsTop + 2
        .rcArea.Right = View.AbsLeft + View.Width - 2
        .rcArea.Bottom = View.AbsTop + View.Height - 2
    End With
    ImmSetCompositionWindow HIMC, Comp
    ImmReleaseContext hWnd, HIMC
End Sub

Private Sub GetCharRange(ByVal sText As String, ByVal nFirst As Long, ByVal nLen As Long, Canvas As Drawing2D.cGraphics, Font As cFont, strFormat As cStringFormat, Bound As Drawing2D.RECTF)
    Dim Range() As Drawing2D.CharacterRange
    Dim Rgn() As cRegion
    Dim sChar As String
    Dim i As Long
    Dim CharRect As Drawing2D.RECTF
    Dim s As String
    
    If nLen > 32 Then Exit Sub
    
    ReDim Range(nLen - 1) As Drawing2D.CharacterRange
    
    For i = 0 To nLen - 1
        Range(i).First = i + nFirst
        Range(i).Length = 1
    Next
    
    strFormat.SetMeasurableCharacterRanges Range
    Rgn = Canvas.MeasureCharacterRanges(sText, Font, Bound, strFormat)
    
    For i = 0 To UBound(Rgn)
        sChar = mID$(sText, i + 1, 1)
        
        Select Case sChar
            Case vbLf
                CharRect = Rgn(i).GetBounds(Canvas)
                CharRect.Right = 1
            Case Else
                CharRect = Rgn(i).GetBounds(Canvas)
        End Select
        With CharRect
            .left = .left + 2
            .Top = .Top + 2
        End With
        CharRange.Add CharRect
    Next
End Sub

