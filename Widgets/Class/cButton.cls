VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButton"
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

Private mCaption    As String
Private mFontName   As String
Private mFontSize   As Single
Private mStyle      As FontStyle
Private mAccent     As Boolean

Private bFocus      As Boolean
Private bDown       As Boolean
Private bHover      As Boolean

Private ThemeColor  As ThemeColorType

Private Sub Class_Initialize()
    mCaption = "Button"
    mFontName = "宋体"
    mFontSize = 14
    mStyle = FontStyleRegular
    
    SetTheme Preset
End Sub

Public Sub Create(Parent As Object, ByVal sCaption As String, ByVal nLeft As Integer, ByVal nTop As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer)
Attribute Create.VB_Description = "创建控件"
    Set View = Parent.CreateView(nLeft, nTop, nWidth, nHeight)
    mCaption = sCaption
    View.Name = sCaption
End Sub

Public Sub BindView(DstView As cView)
Attribute BindView.VB_Description = "绑定已存在的View"
    Dim V As cView
    
    If DstView Is View Then Exit Sub
    
    Set V = View
    Set View = Nothing
    Set View = DstView
    
    bFocus = View.IsFocus
    bHover = False
    bDown = False
    mCaption = "Button"
    
    V.Refresh
    View.Refresh
End Sub

Public Function GetView() As cView
Attribute GetView.VB_Description = "获取控件View"
    Set GetView = View
End Function

'=========================================================================================
'隐藏函数
Public Sub SetAttr(ByVal AttrName As String, ByVal Value As String)
Attribute SetAttr.VB_Description = "设置控件属性"
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
        Case LCase("Caption")
            Me.Caption = Value
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
Attribute Name.VB_Description = "返回/设置对象名称"
Public Property Let Name(ByVal sName As String): View.Name = Trim(sName): End Property

Public Property Get left() As Integer: left = View.left: End Property
Attribute left.VB_Description = "返回/设置对象左边距"
Public Property Let left(ByVal nLeft As Integer)
    If View.left = nLeft Then Exit Property
    View.left = nLeft
End Property

Public Property Get Top() As Integer: Top = View.Top: End Property
Attribute Top.VB_Description = "返回/设置对象顶边距"
Public Property Let Top(ByVal nTop As Integer)
    If View.Top = nTop Then Exit Property
    View.Top = nTop
End Property

Public Property Get Width() As Integer: Width = View.Width: End Property
Attribute Width.VB_Description = "返回/设置对象宽度"
Public Property Let Width(ByVal nWidth As Integer)
    If View.Width = nWidth Then Exit Property
    View.Width = nWidth
End Property

Public Property Get Height() As Integer: Height = View.Height: End Property
Attribute Height.VB_Description = "返回/设置对象高度"
Public Property Let Height(ByVal nHeight As Integer)
    If View.Height = nHeight Then Exit Property
    View.Height = nHeight
End Property

Public Property Get Tag() As String: Tag = View.Tag: End Property
Attribute Tag.VB_Description = "返回/设置对象附带数据"
Public Property Let Tag(ByVal sTag As String): View.Tag = sTag: End Property

Public Property Get Enabled() As Boolean: Enabled = View.Enabled: End Property
Attribute Enabled.VB_Description = "返回/设置对象是否启用"
Public Property Let Enabled(ByVal bEnabled As Boolean)
    If View.Enabled = bEnabled Then Exit Property
    View.Enabled = bEnabled
    View.Refresh
End Property

Public Property Get Visible() As Boolean: Visible = View.Visible: End Property
Attribute Visible.VB_Description = "返回/设置对象可见性"
Public Property Let Visible(ByVal bVisible As Boolean)
    If View.Visible = bVisible Then Exit Property
    View.Visible = bVisible
End Property

Public Property Get Opacity() As Single: Opacity = View.Opacity: End Property
Attribute Opacity.VB_Description = "返回/设置对象透明度"
Public Property Let Opacity(ByVal fOpacity As Single)
    If View.Opacity = fOpacity Then Exit Property
    View.Opacity = fOpacity
End Property

Public Property Get ZOrder() As Long: ZOrder = View.ZOrder: End Property
Attribute ZOrder.VB_Description = "返回/设置对象Z序"
Public Property Let ZOrder(ByVal nZ As Long)
    If View.ZOrder = nZ Then Exit Property
    View.ZOrder = nZ
End Property

Public Sub Move(ByVal nLeft As Integer, Optional nTop, Optional nWidth, Optional nHeight)
Attribute Move.VB_Description = "移动控件"
    Dim T As Integer, W As Integer, H As Integer
    
    T = IIf(IsMissing(nTop), View.Top, CInt(nTop))
    W = IIf(IsMissing(nWidth), View.Width, CInt(nWidth))
    H = IIf(IsMissing(nHeight), View.Height, CInt(nHeight))
    
    View.Move nLeft, T, W, H
End Sub

Public Property Get BackColor() As Long: BackColor = ThemeColor.BKColor: End Property
Attribute BackColor.VB_Description = "返回/设置对象背景色"
Public Property Let BackColor(ByVal nColor As Long)
    If nColor = ThemeColor.BKColor Then Exit Property
    
    ThemeColor.BKColor = nColor
    ThemeColor.BKDrakColor = cColor.SetBrightness(nColor, -0.1)
    ThemeColor.BKLightColor = cColor.SetBrightness(nColor, 0.1)
    
    View.Refresh
End Property

Public Property Get ForeColor() As Long: ForeColor = ThemeColor.TextColor: End Property
Attribute ForeColor.VB_Description = "返回/设置对象前景色"
Public Property Let ForeColor(ByVal nColor As Long)
    If nColor = ThemeColor.TextColor Then Exit Property
    
    ThemeColor.TextColor = nColor
    View.Refresh
End Property

Public Sub Refresh(): View.Refresh: End Sub
Attribute Refresh.VB_Description = "刷新控件显示"

Public Sub SetTheme(NewTheme As PresetThemeEnum)
Attribute SetTheme.VB_Description = "设置一个预设主题"
    Select Case NewTheme
        Case PresetThemeEnum.DrakTheme
            SetDarkTheme ThemeColor
        Case PresetThemeEnum.LightTheme
            SetLightTheme ThemeColor
    End Select
End Sub

'=========================================================================================
Public Property Get Caption() As String: Caption = mCaption: End Property
Attribute Caption.VB_Description = "返回/设置对象标题"
Public Property Let Caption(ByVal sCaption As String)
    If mCaption = sCaption Then Exit Property
    mCaption = sCaption
    View.Refresh
End Property

Public Property Get FontName() As String: FontName = mFontName: End Property
Attribute FontName.VB_Description = "返回/设置字体名称"
Public Property Let FontName(ByVal sFontName As String)
    If mFontName = sFontName Then Exit Property
    mFontName = sFontName
    View.Refresh
End Property

Public Property Get FontSize() As Single: FontSize = mFontSize: End Property
Attribute FontSize.VB_Description = "返回/设置字体大小"
Public Property Let FontSize(ByVal nSize As Single)
    If mFontSize = nSize Then Exit Property
    mFontSize = nSize
    View.Refresh
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "返回/设置文本是否显示为粗体"
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
Attribute FontItalic.VB_Description = "返回/设置文本是否显示为斜体"
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
Attribute FontUnderLine.VB_Description = "返回/设置文本是否有下划线"
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
Attribute FontStrikeout.VB_Description = "返回/设置文本是否有删除线"
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
Attribute IsAccent.VB_Description = "返回/设置对象是否着重显示"
Public Property Let IsAccent(ByVal bAccent As Boolean)
    If mAccent = bAccent Then Exit Property
    mAccent = bAccent
    View.Refresh
End Property

Public Property Get AccentColor() As Long: AccentColor = ThemeColor.AccentColor: End Property
Attribute AccentColor.VB_Description = "返回/设置对象着重显示颜色"
Public Property Let AccentColor(ByVal nColor As Long)
    If ThemeColor.AccentColor = nColor Then Exit Property
    ThemeColor.AccentColor = nColor
    ThemeColor.AccentDarkColor = cColor.SetBrightness(nColor, -0.2)
    ThemeColor.AccentLightColor = cColor.SetBrightness(nColor, 0.15)
    View.Refresh
End Property

Public Property Get BorderColor() As Long: BorderColor = ThemeColor.BorderColor: End Property
Attribute BorderColor.VB_Description = "返回/设置对象边框颜色"
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
Private Sub View_LostFocus()
    RaiseEvent LostFocus
    bFocus = False
    View.Refresh
End Sub

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
    Dim Brush As New cSolidBrush
    
    Dim Fam As New cFontFamily
    Dim Font As New cFont
    Dim strFormat As New cStringFormat
    
    Dim Bound As RECTF
    
    Fam.FromName mFontName
    Font.FromFamily Fam, mFontSize, mStyle, UnitPixel
    strFormat.Align = StringAlignmentCenter
    strFormat.LineAlign = StringAlignmentCenter
    
    If View.Enabled Then
        If mAccent Then
            If bHover Then
                Canvas.Clear ThemeColor.AccentLightColor
                If bDown Then
                    Canvas.Clear ThemeColor.AccentDarkColor
                End If
            Else
                Canvas.Clear ThemeColor.AccentColor
            End If
            
            Pen.Color = ThemeColor.AccentColor
            Brush.Color = cColor.FromARGB(255, 240, 240, 240)
        Else
            If bHover Then
                Canvas.Clear ThemeColor.BKLightColor
                If bDown Then
                    Canvas.Clear ThemeColor.BKDrakColor
                End If
            Else
                Canvas.Clear ThemeColor.BKColor
            End If
            
            If bFocus Then
                Pen.Color = ThemeColor.AccentColor
            Else
                Pen.Color = ThemeColor.BorderColor
            End If
            
            Brush.Color = ThemeColor.TextColor
        End If
    Else
        If mAccent Then
            Canvas.Clear ThemeColor.AccentLightColor
            Pen.Color = ThemeColor.AccentLightColor
            Brush.Color = cColor.SetBrightness(ThemeColor.AccentLightColor, 0.2)
        Else
            Canvas.Clear ThemeColor.BKDrakColor
            Pen.Color = ThemeColor.BorderDisEnableColor
            Brush.Color = ThemeColor.BorderDisEnableColor
        End If
    End If
    
    Bound = NewRectF(0, 0, View.Width - 1, View.Height - 1)
    
    Canvas.SetTextRenderingHint TextRenderingHintAntiAliasGridFit
    Canvas.DrawRectangle Pen, 0, 0, View.Width - 1, View.Height - 1
    Canvas.DrawString mCaption, Font, Bound, strFormat, Brush.GetBaseBrush
End Sub
