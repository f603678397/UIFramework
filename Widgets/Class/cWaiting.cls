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

Private Points(4)   As Single

Private Sub Class_Initialize()
    Set Timer = New cTimer
    SetTheme Preset
End Sub

Public Sub Create(Parent As Object, ByVal sCaption As String, ByVal nLeft As Integer, ByVal nTop As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer)
    Dim nhWnd       As Long
    
    nWidth = IIf(nWidth < 20, 20, nWidth)
    nHeight = IIf(nHeight < 20, 20, nHeight)
    
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
    nWidth = IIf(nWidth < 20, 20, nWidth)
    If View.Width = nWidth Then Exit Property
    View.Width = nWidth
End Property

Public Property Get Height() As Integer: Height = View.Height: End Property
Public Property Let Height(ByVal nHeight As Integer)
    nHeight = IIf(nHeight < 20, 20, nHeight)
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
    Static Angle As Single
    Static B As Byte
    Dim R As Single
    Dim Offset As Single
    Dim i As Integer
    
    If Angle < 135 Or Angle > 225 Then
        R = Angle / 180 * PI
        Offset = 20 * Cos(R) + 20
    Else
        Offset = 3
    End If
    
    If Angle < 360 Then
        Angle = Angle + Offset
    Else
        Angle = 0
        B = IIf(B < 2, B + 1, 0)
    End If
    
    For i = UBound(Points) To 1 Step -1
        Points(i) = Points(i - 1) - 10
    Next
    
    If B <> 2 Then
        Points(0) = Angle
    Else
        Points(0) = -1
    End If

    View.Refresh
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Brush As New cSolidBrush
    Dim L As Integer, T As Integer, N As Integer
    Dim cX As Single, cY As Single, cR As Single
    Dim i As Integer
    Dim R As Single
    
    Brush.Color = ThemeColor.AccentColor
    
    N = IIf(View.Width < View.Height, View.Width, View.Height) / 2
    N = N - N / 5
    cR = N / 5
    L = View.Width / 2
    T = View.Height / 2
    
    
    Canvas.SetSmoothingMode SmoothingModeAntiAlias8x8
    
    For i = 0 To UBound(Points)
        If Points(i) > 0 Then
            R = (Points(i) + 90) / 180 * PI
            
            cX = L + Cos(R) * N
            cY = T + Sin(R) * N
            Canvas.FillEllipse Brush.GetBaseBrush, cX - cR / 2, cY - cR / 2, cR, cR
        End If
    Next
End Sub
