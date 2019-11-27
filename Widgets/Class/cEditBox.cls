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
Public Event KeyDown(ByVal KeyCode As Integer)
Public Event KeyPress(ByVal KeyAscii As Long)
Public Event KeyUp(ByVal KeyCode As Integer)
Public Event TextChanged()

Private mText       As String
Private mFontName   As String
Private mFontSize   As Single
Private mStyle      As FontStyle
Private mAccent     As Boolean
Private mAutoWrap   As Boolean

Private bFocus      As Boolean
Private bDown       As Boolean
Private bChanged    As Boolean

Private ThemeColor  As ThemeColorType

Private CharRange   As New cArrayList
Private InputPos    As Drawing2D.RECTF
Private CharIndex   As Long

Private MI_IBEAM    As Long

Private OffsetX     As Long
Private OffsetY     As Long


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
    Dim i As Long
    
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

Public Property Get AutoWrap() As Boolean: AutoWrap = mAutoWrap: End Property
Public Property Let AutoWrap(ByVal bAuto As Boolean)
    If bAuto = mAutoWrap Then Exit Property
    mAutoWrap = bAuto
    bChanged = True
    View.Refresh
End Property

Private Sub View_Click(): RaiseEvent Click: End Sub
Private Sub View_DblClick(): RaiseEvent DblClick: End Sub

Private Sub View_GotFocus()
    RaiseEvent GotFocus
    bFocus = True
    View.Refresh
End Sub

Private Sub View_KeyPress(ByVal KeyAscii As Long)
'
End Sub

Private Sub View_LostFocus()
    RaiseEvent LostFocus
    bFocus = False
    View.Refresh
End Sub

Private Sub View_MouseHover()
    RaiseEvent MouseHover
End Sub
Private Sub View_MouseLeave()
    RaiseEvent MouseLeave
End Sub

Private Sub View_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim Bound As Drawing2D.RECTF
    Dim W As Single
    Dim i As Long
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    bDown = True

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
    Dim txtPen      As New cPen
    Dim Brush       As New cSolidBrush

    If View.Enabled Then
        If mAccent Or bFocus Then
            Pen.Color = ThemeColor.AccentColor
        Else
            Pen.Color = ThemeColor.BorderColor
        End If
        Brush.Color = ThemeColor.TextColor
        txtPen.Color = cColor.Black
        Canvas.Clear cColor.White
    Else
        Pen.Color = ThemeColor.BorderDisEnableColor
        Brush.Color = ThemeColor.BorderDisEnableColor
        txtPen.Color = ThemeColor.BorderDisEnableColor
        Canvas.Clear ThemeColor.BKColor
    End If
    
    If bChanged Then
        DrawText Brush.GetBaseBrush
    End If
    
    Canvas.DrawImage Bitmap.ToImage, 2 + OffsetX, 2 + OffsetY
    Canvas.DrawRectangle Pen, 0, 0, View.Width - 1, View.Height - 1
End Sub

Private Sub DrawText(Brush As cBrush)
    Dim Fam             As New cFontFamily
    Dim Font            As New cFont
    Dim strFormat       As New cStringFormat
    
    Dim Canvas          As cGraphics
    Dim BmpHeight       As Long
    
    Dim sText           As String
    Dim Bound           As Drawing2D.RECTF
    
    Dim nFirst          As Long
    Dim nLen            As Long
    Dim txtLen          As Long
    
    sText = Replace(mText, vbCrLf, vbLf)
    
    Fam.FromName mFontName
    Font.FromFamily Fam, mFontSize, mStyle, UnitPixel
    strFormat.Flags = StringFormatFlagsNoClip Or StringFormatFlagsMeasureTrailingSpaces
    If Not mAutoWrap Then strFormat.Flags = strFormat.Flags Or StringFormatFlagsNoWrap
    
    BmpHeight = GetBitmapHeight(Bitmap.GetGraphics, Font, strFormat)
    Bitmap.CreateEmptyBitmap View.Width - 4, BmpHeight
    Set Canvas = Bitmap.GetGraphics
    
    Bound = NewRectF(0, 0, Bitmap.Width, Bitmap.Height)

    txtLen = Len(mText)
    CharRange.Clear
    
    If txtLen < 33 Then
        nLen = txtLen
        GetCharRange nFirst, nLen, Canvas, Font, strFormat, Bound
    Else
        Do
            GetCharRange nFirst, nLen, Canvas, Font, strFormat, Bound
            nFirst = nFirst + nLen
            If txtLen - nFirst > 32 Then
                nLen = 32
            Else
                nLen = txtLen - nFirst
            End If
        Loop While nFirst < txtLen
    End If
    
    Canvas.DrawString mText, Font, Bound, strFormat, Brush
End Sub

Private Function GetBitmapHeight(ByVal sText As String, Canvas As Drawing2D.cGraphics, Font As cFont, strFormat As cStringFormat) As Long
    Dim Bound       As Drawing2D.RECTF
    
    Bound.Right = View.Width - 4
    
    Bound = Canvas.MeasureString(sText, Font, Bound, strFormat)
    GetBitmapHeight = Bound.Bottom
End Function

Private Function GetInputPosition()
    InputPos = CharRange.GetValue(CharIndex)
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
    Rgn = Canvas.MeasureCharacterRanges(mText, Font, Bound, strFormat)
    
    For i = 0 To UBound(Rgn)
        sChar = mID$(mText, i + 1, 1)
        
        Select Case sChar
            Case vbLf
            
            Case Else
                CharRect = Rgn(i).GetBounds(Canvas)
        End Select
        CharRange.Add CharRect
    Next
End Sub

