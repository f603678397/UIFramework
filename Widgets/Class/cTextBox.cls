VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents View        As cView
Attribute View.VB_VarHelpID = -1

Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseHover()
Public Event MouseLeave()
Public Event GotFocus()
Public Event LostFocus()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(ByVal KeyCode As Integer)
Public Event KeyPress(ByVal KeyAscii As Long)
Public Event KeyUp(ByVal KeyCode As Integer)
Public Event TextChanged()

Private mText       As String
Private mFontName   As String
Private mFontSize   As Single
Private mStyle      As FontStyle
Private mAccent     As Boolean

Private bFocus      As Boolean
Private bDown       As Boolean
Private bChanged    As Boolean

Private ThemeColor  As ThemeColorType
Private CharRange   As New cArrayList

Private MI_IBEAM    As Long
Private CharIndex   As Long

Private Sub Class_Initialize()
    mText = "TextBox"
    mFontName = "宋体"
    mFontSize = 14
    mStyle = FontStyleRegular
    bChanged = True
    
    MI_IBEAM = LoadCursorWL(ByVal 0&, IDC_IBEAM)
    
    SetTheme Preset
End Sub

Private Sub Class_Terminate()
    DestroyCursor MI_IBEAM
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
Attribute SetAttr.VB_MemberFlags = "40"
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
    View.Refresh
End Property

Public Property Get FontSize() As Single: FontSize = mFontSize: End Property
Public Property Let FontSize(ByVal nSize As Single)
    If mFontSize = nSize Then Exit Property
    mFontSize = nSize
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

Private Sub View_Click(): RaiseEvent Click: End Sub
Private Sub View_DblClick(): RaiseEvent DblClick: End Sub

Private Sub View_GotFocus()
    RaiseEvent GotFocus
    bFocus = True
    View.Refresh
End Sub

Private Sub View_KeyDown(ByVal KeyCode As Integer)
    Dim sChar As String
    Select Case KeyCode
        Case vbKeyLeft
            If CharIndex > 0 Then CharIndex = CharIndex - 1
            View.Refresh
        Case vbKeyRight
            If CharIndex < Len(mText) Then CharIndex = CharIndex + 1
            View.Refresh
    End Select
End Sub

Private Sub View_KeyPress(ByVal KeyAscii As Long)
    Dim sChar As String
    Dim sText As String
    Dim s1 As String, s2 As String
    
    Select Case KeyAscii
        Case vbKeyReturn
            sChar = vbCrLf
            s1 = left$(mText, CharIndex)
            s2 = Right$(mText, Len(mText) - CharIndex)
            sText = s1 & sChar & s2
            CharIndex = CharIndex + 2
        Case vbKeyTab
            sChar = "  "
            s1 = left$(mText, CharIndex)
            s2 = Right$(mText, Len(mText) - CharIndex)
            sText = s1 & sChar & s2
            CharIndex = CharIndex + 2
        Case vbKeySpace
            sChar = " "
            s1 = left$(mText, CharIndex)
            s2 = Right$(mText, Len(mText) - CharIndex)
            sText = s1 & sChar & s2
            CharIndex = CharIndex + 1
        Case vbKeyBack
            If CharIndex > 0 Then
                If CharIndex > 1 Then
                    sChar = mID$(mText, CharIndex - 1, 2)
                Else
                    sChar = "*"
                End If
                
                If sChar = vbCrLf Then
                    s1 = left$(mText, CharIndex - 2)
                    s2 = Right$(mText, Len(mText) - CharIndex)
                    sText = s1 & s2
                    CharIndex = CharIndex - 2
                Else
                    s1 = left$(mText, CharIndex - 1)
                    s2 = Right$(mText, Len(mText) - CharIndex)
                    sText = s1 & s2
                    CharIndex = CharIndex - 1
                End If
            Else
                Exit Sub
            End If
        Case Else
            If KeyAscii < 32 Then
                Exit Sub
            End If
            sChar = Chr$(KeyAscii)
            s1 = left$(mText, CharIndex)
            s2 = Right$(mText, Len(mText) - CharIndex)
            sText = s1 & sChar & s2
            CharIndex = CharIndex + 1
    End Select

    Me.Text = sText
End Sub

Private Sub View_LostFocus()
    RaiseEvent LostFocus
    bFocus = False
    View.Refresh
End Sub

Private Sub View_MouseHover()
    RaiseEvent MouseHover
    SetSystemCursor MI_IBEAM, OCR_NORMAL
End Sub
Private Sub View_MouseLeave()
    RaiseEvent MouseLeave
    SetSystemCursor MI_IBEAM, OCR_NORMAL
End Sub

Private Sub View_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim Bound As Drawing2D.RECTF
    Dim W As Single
    Dim i As Long
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    bDown = True
    
    CharIndex = Len(mText)
    For i = 0 To CharRange.GetCount - 1
        Bound = CharRange.GetValue(i)
        
        If Y >= Bound.Top And Y <= Bound.Top + Bound.Bottom Then
            CharIndex = i
            If mID$(mText, CharIndex + 1, 1) <> vbLf Then CharIndex = i + 1
            If X >= Bound.left And X <= Bound.left + Bound.Right Then
                W = Bound.Right
                If X - Bound.left < W / 2 Then
                    CharIndex = i
                Else
                    CharIndex = i + 1
                End If
                Exit For
            End If
        End If
    Next
    
    View.Refresh
End Sub

Private Sub View_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub View_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    bDown = False
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen         As New cPen
    Dim Brush       As New cSolidBrush
    
    Dim Fam         As New cFontFamily
    Dim Font        As New cFont
    Dim strFormat   As New cStringFormat
    
    Dim i           As Long
    Dim Bound       As Drawing2D.RECTF
    Dim InputPos    As Drawing2D.RECTF
    Dim CharRect    As Drawing2D.RECTF
    Dim sChar       As String
    Dim nFirst      As Long
    Dim nLen        As Long
    Dim HIMC        As Long
    Dim Comp        As COMPOSITIONFORM
    Dim hWnd        As Long
    
    Fam.FromName mFontName
    Font.FromFamily Fam, mFontSize, mStyle, UnitPixel
    strFormat.Flags = StringFormatFlagsNoClip
    
    Bound = NewRectF(2, 2, View.Width - 2, View.Height - 2)
    
    If bChanged Then
        CharRange.Clear
        If Len(mText) > 0 Then
            If Len(mText) < 33 Then
                GetCharRange 0, Len(mText), Canvas, Font, strFormat, Bound
            Else
                nFirst = 0
                nLen = 32
                Do
                    GetCharRange nFirst, nLen, Canvas, Font, strFormat, Bound
                    nFirst = nFirst + nLen
                    If Len(mText) - nFirst > 32 Then
                        nLen = 32
                    Else
                        nLen = Len(mText) - nFirst
                    End If
                Loop While nFirst < Len(mText)
            End If
        End If
        bChanged = False
    End If
    
    If CharRange.GetCount > 0 Then
        If CharIndex < CharRange.GetCount Then
            CharRect = CharRange.GetValue(CharIndex)
            InputPos.left = CharRect.left
        Else
            CharRect = CharRange.GetValue(CharRange.GetCount - 1)
            InputPos.left = CharRect.left + CharRect.Right
        End If
        InputPos.Top = CharRect.Top
        InputPos.Right = InputPos.left
        InputPos.Bottom = CharRect.Top + CharRect.Bottom
    Else
        InputPos.left = 2
        InputPos.Top = 2
        InputPos.Right = 2
        InputPos.Bottom = mFontSize
    End If
    
    hWnd = View.GetLayout.GetActivity.hWnd
    HIMC = ImmGetContext(hWnd)
    With Comp
        .dwStyle = CFS_POINT Or CFS_RECT
        .ptCurrentPos.X = View.AbsLeft + InputPos.left
        .ptCurrentPos.Y = View.AbsTop + InputPos.Top
        .rcArea.left = View.AbsLeft + 2
        .rcArea.Top = View.AbsTop + 2
        .rcArea.Right = View.AbsLeft + View.Width - 2
        .rcArea.Bottom = View.AbsTop + View.Height - 2
    End With
    ImmSetCompositionWindow HIMC, Comp
    ImmReleaseContext hWnd, HIMC

    If View.Enabled Then
        If mAccent Or bFocus Then
            Pen.Color = ThemeColor.AccentColor
        Else
            Pen.Color = ThemeColor.BorderColor
        End If
        Brush.Color = ThemeColor.TextColor
    Else
        Pen.Color = ThemeColor.BorderDisEnableColor
        Brush.Color = ThemeColor.BorderDisEnableColor
    End If
    
    Canvas.Clear cColor.White
    Canvas.SetTextRenderingHint TextRenderingHintClearTypeGridFit
    Canvas.DrawRectangle Pen, 0, 0, View.Width - 1, View.Height - 1
    Canvas.DrawString mText, Font, Bound, strFormat, Brush.GetBaseBrush
    
    Pen.Color = cColor.Black
    
    If bFocus Then
        With InputPos
                Canvas.DrawLine Pen, .left, .Top, .Right, .Bottom
        End With
    End If
End Sub

Private Sub GetCharRange(ByVal nFirst As Long, ByVal nLen As Long, Canvas As Drawing2D.cGraphics, Font As cFont, strFormat As cStringFormat, Bound As Drawing2D.RECTF)
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
    Rgn = Canvas.MeasureCharacterRanges(mText & "a", Font, Bound, strFormat)
    
    For i = 0 To UBound(Rgn)
        sChar = mID$(mText, i + 1, 1)
        
        If sChar = vbCr Then
            If i > 0 Then
                CharRect = Rgn(i - 1).GetBounds(Canvas)
                CharRect.left = CharRect.left + CharRect.Right
                CharRect.Right = 1
                If CharRect.Bottom = 0 Then CharRect.Bottom = mFontSize
            Else
                CharRect.left = 2
                CharRect.Top = 2
                CharRect.Right = 1
                CharRect.Bottom = mFontSize
            End If
        ElseIf sChar = vbLf Then
            CharRect = Rgn(i - 1).GetBounds(Canvas)
            CharRect.left = 2
            If CharRect.Bottom = 0 Then CharRect.Bottom = mFontSize
            CharRect.Top = CharRect.Top + CharRect.Bottom
        Else
            CharRect = Rgn(i).GetBounds(Canvas)
        End If
        
        CharRange.Add CharRect
    Next
End Sub
