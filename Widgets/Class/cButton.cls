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

Private Sub Class_Initialize()
    mCaption = "Button"
    mFontName = "宋体"
    mFontSize = 14
    mStyle = FontStyleRegular
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
    
    bFocus = View.IsFocus
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

Public Sub Refresh(): View.Refresh: End Sub

'=========================================================================================
Public Property Get Caption() As String: Caption = mCaption: End Property
Public Property Let Caption(ByVal sCaption As String)
    If mCaption = sCaption Then Exit Property
    mCaption = sCaption
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

Public Property Get IsAccent() As Boolean: IsAccent = mAccent: End Property
Public Property Let IsAccent(ByVal bAccent As Boolean)
    If mAccent = bAccent Then Exit Property
    mAccent = bAccent
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
        Else
            If bHover Then
                Canvas.Clear ThemeColor.PrimaryLightColor
                If bDown Then
                    Canvas.Clear ThemeColor.PrimaryDrakColor
                End If
            Else
                Canvas.Clear ThemeColor.PrimaryColor
            End If
            
            If bFocus Then
                Pen.Color = ThemeColor.AccentColor
            Else
                Pen.Color = ThemeColor.BorderColor
            End If
        End If
        
        Brush.Color = ThemeColor.TextColor
    Else
        If mAccent Then
            Canvas.Clear ThemeColor.AccentDarkColor
            Pen.Color = ThemeColor.AccentDarkColor
            Brush.Color = ThemeColor.AccentLightColor
        Else
            Canvas.Clear ThemeColor.PrimaryDrakColor
            Pen.Color = ThemeColor.BorderColor
            Brush.Color = ThemeColor.PrimaryLightColor
        End If
    End If
    
    Bound = NewRectF(0, 0, View.Width - 1, View.Height - 1)
    
    Canvas.SetTextRenderingHint TextRenderingHintClearTypeGridFit
    Canvas.DrawRectangle Pen, 0, 0, View.Width - 1, View.Height - 1
    Canvas.DrawString mCaption, Font, Bound, strFormat, Brush.GetBaseBrush
End Sub
