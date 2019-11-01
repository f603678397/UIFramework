VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents View         As cView
Attribute View.VB_VarHelpID = -1
Private WithEvents Timer        As cTimer
Attribute Timer.VB_VarHelpID = -1

Private ThemeColor  As ThemeColorType

Private StartAngle  As Single
Private EndAngle    As Single

Private Sub Class_Initialize()
    Set Timer = New cTimer
    
    StartAngle = 30
    EndAngle = -30
    
    SetTheme Preset
End Sub

Public Sub Create(Parent As Object, ByVal sCaption As String, ByVal nLeft As Integer, ByVal nTop As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer)
    Dim nhWnd       As Long
    Set View = Parent.CreateView(nLeft, nTop, nWidth, nHeight)
    nhWnd = View.GetLayout.GetActivity.hWnd
    Timer.Create nhWnd
    Timer.Interval = 33
    View.Name = sCaption
End Sub

Public Sub BindView(DstView As cView)
    Dim V           As cView
    Dim nhWnd       As Long
    
    If DstView Is View Then Exit Sub
    
    Timer.Release
    
    Set V = View
    Set View = Nothing
    Set View = DstView
    
    nhWnd = View.GetLayout.GetActivity.hWnd
    Timer.Create nhWnd
    Timer.Enabled = 33
    
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
        Case LCase("Visible")
            Me.Visible = CBool(Value)
        Case LCase("Opacity")
            Me.Opacity = CSng(Value)
    '-----------------------------------------
        Case LCase("AccentColor")
            Me.AccentColor = CLng(Value)
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

Public Property Get Visible() As Boolean: Visible = View.Visible: End Property
Public Property Let Visible(ByVal bVisible As Boolean)
    If View.Visible = bVisible Then Exit Property
    View.Visible = bVisible
    Timer.Enabled = bVisible
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

Public Sub SetTheme(NewTheme As PresetThemeEnum)
    Select Case NewTheme
        Case PresetThemeEnum.DrakTheme
            SetDarkTheme ThemeColor
        Case PresetThemeEnum.LightTheme
            SetLightTheme ThemeColor
    End Select
End Sub

'=========================================================================================
Public Property Get AccentColor() As Long: AccentColor = ThemeColor.AccentColor: End Property
Public Property Let AccentColor(ByVal nColor As Long)
    If ThemeColor.AccentColor = nColor Then Exit Property
    ThemeColor.AccentColor = nColor
    ThemeColor.AccentDarkColor = cColor.SetBrightness(nColor, -0.2)
    ThemeColor.AccentLightColor = cColor.SetBrightness(nColor, 0.15)
    View.Refresh
End Property

Private Sub Timer_onTime()
    Static A As Single
    Dim R As Single
    
    R = StartAngle / 720 * PI
    StartAngle = StartAngle + 30 * Sin(R) + 1.5
    StartAngle = IIf(StartAngle < 720, StartAngle, StartAngle - 720)
    
    EndAngle = -135 * Sin(R) - 30
    
    View.Refresh
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen As New cPen
    Dim L As Integer, T As Integer, W As Integer
    
    Pen.Width = 3
    Pen.Color = ThemeColor.AccentColor
    
    W = IIf(View.Width < View.Height, View.Width - 5, View.Height - 5)
    L = (View.Width - W) / 2
    T = (View.Height - W) / 2
    
    Canvas.SetSmoothingMode SmoothingModeAntiAlias
    Canvas.SetTranslate -View.Width / 2, -View.Height / 2, MatrixOrderPrepend
    Canvas.SetRotate -90, MatrixOrderAppend
    Canvas.SetTranslate View.Width / 2, View.Height / 2, MatrixOrderAppend
    Canvas.DrawArc Pen, L, T, W, W, StartAngle, EndAngle
    Canvas.ResetWorldTransform
End Sub
