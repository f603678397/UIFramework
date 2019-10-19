VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOption"
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
Public Event ValueChanged(ByVal ByUser As Boolean)

Private mCaption    As String
Private mValue      As Boolean
Private mFontName   As String
Private mFontSize   As Single
Private mStyle      As FontStyle

Private bDown       As Boolean
Private bHover      As Boolean
Private mID         As Long


Private ThemeColor  As ThemeColorType

Private Sub Class_Initialize()
    mCaption = "Button"
    mFontName = "宋体"
    mFontSize = 14
    mStyle = FontStyleRegular
    
    SetTheme Preset
    mID = ObjPtr(Me)
End Sub

Public Sub Create(Parent As Object, ByVal sCaption As String, ByVal nLeft As Integer, ByVal nTop As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer)
    Set View = Parent.CreateView(nLeft, nTop, nWidth, nHeight)
    mCaption = sCaption
    View.Name = sCaption
End Sub

Public Sub BindView(DstView As cView)
    Dim V As cView
    
    If DstView Is View Then Exit Sub
    
    Set V = View
    Set View = Nothing
    Set View = DstView

    bHover = False
    bDown = False
    
    V.Refresh
    View.Refresh
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

Public Property Get BackColor() As Long: BackColor = ThemeColor.PrimaryColor: End Property
Public Property Let BackColor(ByVal nColor As Long)
    If nColor = ThemeColor.PrimaryColor Then Exit Property
    
    ThemeColor.PrimaryColor = nColor
    ThemeColor.PrimaryDrakColor = cColor.SetBrightness(nColor, -0.1)
    ThemeColor.PrimaryLightColor = cColor.SetBrightness(nColor, 0.1)
    
    View.Refresh
End Property

Public Property Get ForeColor() As Long: ForeColor = ThemeColor.TextColor: End Property
Public Property Let ForeColor(ByVal nColor As Long)
    If nColor = ThemeColor.TextColor Then Exit Property
    
    ThemeColor.TextColor = nColor
    View.Refresh
End Property

Public Property Get FocusColor() As Long: FocusColor = ThemeColor.AccentColor: End Property
Public Property Let FocusColor(ByVal nColor As Long)
    If nColor = ThemeColor.AccentColor Then Exit Property
    
    ThemeColor.AccentColor = nColor
    ThemeColor.AccentDarkColor = cColor.SetBrightness(nColor, -0.1)
    ThemeColor.AccentLightColor = cColor.SetBrightness(nColor, 0.1)
    
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
    
    ThemeColor.PrimaryColor = cColor.Transparent
End Sub

'=========================================================================================
Public Property Get Caption() As String: Caption = mCaption: End Property
Public Property Let Caption(ByVal sCaption As String)
    If mCaption = sCaption Then Exit Property
    mCaption = sCaption
    View.Refresh
End Property

Public Property Get Value() As Boolean: Value = mValue: End Property
Public Property Let Value(ByVal bValue As Boolean)
    If mValue = bValue Then Exit Property
    mValue = bValue
    ChangeValue False
    View.Refresh
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

Private Sub ChangeValue(ByVal ByUser As Boolean)
    View.SendMsgToSameLevel VM_OPTION, mID, 0
    RaiseEvent ValueChanged(ByUser)
End Sub

Private Sub View_Click()
    RaiseEvent Click
    If mValue = False Then
        mValue = True
        ChangeValue True
        View.Refresh
    End If
End Sub
Private Sub View_DblClick(): RaiseEvent DblClick: End Sub

Private Sub View_GotFocus(): RaiseEvent GotFocus: End Sub

Private Sub View_HandleMsg(ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    If nMsg <> VM_OPTION Then Exit Sub
    If wParam = mID Or mValue = False Then Exit Sub
    mValue = False
    RaiseEvent ValueChanged(False)
    View.Refresh
End Sub

Private Sub View_LostFocus(): RaiseEvent LostFocus: End Sub

Private Sub View_MouseHover()
    RaiseEvent MouseHover
    bHover = True
    View.Refresh
End Sub

Private Sub View_MouseLeave()
    RaiseEvent MouseLeave
    bHover = False
    View.Refresh
End Sub

Private Sub View_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    bDown = True
    View.Refresh
End Sub

Private Sub View_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub View_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    bDown = False
    View.Refresh
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen As New cPen
    Dim Brush2 As New cSolidBrush
    Dim Brush As New cSolidBrush
    
    Dim Fam As New cFontFamily
    Dim Font As New cFont
    Dim strFormat As New cStringFormat
    
    Dim Bound As RECTF
    
    Fam.FromName mFontName
    Font.FromFamily Fam, mFontSize, mStyle, UnitPixel
    strFormat.Align = StringAlignmentNear
    strFormat.LineAlign = StringAlignmentCenter
    
    If View.Enabled Then
        If bHover Then
            If bDown Then
                Pen.Color = ThemeColor.AccentDarkColor
            Else
                Pen.Color = ThemeColor.AccentLightColor
            End If
        Else
            Pen.Color = ThemeColor.AccentColor
        End If
        Brush.Color = ThemeColor.TextColor
    Else
        Brush.Color = ThemeColor.PrimaryLightColor
        Pen.Color = ThemeColor.PrimaryLightColor
    End If
    
    Brush2.Color = Pen.Color
    Pen.Width = 2
    
    Bound = NewRectF(20, 1, View.Width - 21, View.Height - 1)
    
    Canvas.SetTextRenderingHint TextRenderingHintAntiAliasGridFit
    Canvas.SetSmoothingMode SmoothingModeAntiAlias
    Canvas.Clear ThemeColor.PrimaryColor
    
    Canvas.DrawEllipse Pen, 4, (View.Height - 1 - 10) / 2, 10, 10
    If mValue Then
        Canvas.FillEllipse Brush2.GetBaseBrush, 6, (View.Height - 1 - 6) / 2, 6, 6
    End If
    Canvas.DrawString mCaption, Font, Bound, strFormat, Brush.GetBaseBrush
End Sub
