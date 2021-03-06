VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSeekBar"
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
Public Event ProgressChanged(ByVal ByUser As Boolean)

Private mValue          As Long
Private mMax            As Long
Private bDown           As Boolean
Private bHover          As Boolean

Private ThemeColor  As ThemeColorType

Private Sub Class_Initialize()
    mMax = 100
    SetTheme Preset
End Sub

Public Sub Create(Parent As Object, ByVal sName As String, ByVal nLeft As Integer, ByVal nTop As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer)
    Set View = Parent.CreateView(nLeft, nTop, nWidth, nHeight)
    View.Name = sName
End Sub

Public Sub BindView(DstView As cView)
    Dim V As cView
    
    If DstView Is View Then Exit Sub
    
    Set V = View
    Set View = Nothing
    Set View = DstView
    
    bDown = False
    
    V.Refresh
    View.Refresh
End Sub

Public Function GetView() As cView
    Set GetView = View
End Function

'=========================================================================================
'���غ���
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
        Case LCase("Value")
            Me.Value = CLng(Value)
        Case LCase("Max")
            Me.Max = CLng(Value)
        Case LCase("BackColor")
            Me.BackColor = CLng(Value)
        Case LCase("ForeColor")
            Me.ForeColor = CLng(Value)
        Case LCase("AccentColor")
            Me.AccentColor = CLng(Value)
        Case LCase("BorderColor")
            Me.BorderColor = CLng(Value)
    '-----------------------------------------
    End Select
End Sub


'=========================================================================================
'ͨ�ô���

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

Public Property Get ForeColor() As Long: ForeColor = ThemeColor.AccentColor: End Property
Public Property Let ForeColor(ByVal nColor As Long): Me.AccentColor = nColor: End Property

Public Sub Refresh(): View.Refresh: End Sub

Public Sub SetTheme(NewTheme As PresetThemeEnum)
    Select Case NewTheme
        Case PresetThemeEnum.DrakTheme
            SetDarkTheme ThemeColor
        Case PresetThemeEnum.LightTheme
            SetLightTheme ThemeColor
            ThemeColor.BKColor = cColor.FromARGB(255, 210, 210, 210)
            ThemeColor.BKDrakColor = cColor.FromARGB(255, 220, 220, 220)
    End Select
    ThemeColor.BorderColor = cColor.Transparent
    ThemeColor.BorderDisEnableColor = cColor.Transparent
End Sub

'=========================================================================================
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

Public Property Get Value() As Long: Value = mValue: End Property
Public Property Let Value(ByVal nValue As Long)
    If mValue = nValue Or bDown Then Exit Property
    mValue = nValue
    RaiseEvent ProgressChanged(False)
    View.Refresh
End Property

Public Property Get Max() As Long: Max = mMax: End Property
Public Property Let Max(ByVal nMax As Long)
    If mMax = nMax Then Exit Property
    If nMax = 0 Then nMax = 1
    mMax = nMax
    View.Refresh
End Property

Private Sub View_Click(): RaiseEvent Click: End Sub
Private Sub View_DblClick(): RaiseEvent DblClick: End Sub

Private Sub View_GotFocus(): RaiseEvent GotFocus: End Sub
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
    Dim P As Single
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = vbLeftButton Then
        bDown = True
        P = CSng(X) / View.Width
        P = IIf(P < 0, 0, IIf(P > 1, 1, P))
        mValue = mMax * P
        RaiseEvent ProgressChanged(True)
        View.Refresh
    End If
End Sub

Private Sub View_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim P As Single
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If bDown Then
        P = CSng(X) / View.Width
        P = IIf(P < 0, 0, IIf(P > 1, 1, P))
        mValue = mMax * P
        RaiseEvent ProgressChanged(True)
        View.Refresh
    End If
End Sub

Private Sub View_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If bDown Then
        bDown = False
        View.Refresh
    End If
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen1    As New cPen
    Dim Pen2    As New cPen
    Dim Brush   As New cSolidBrush
    
    Dim P As Single
    Dim L As Single, T As Single
    
    P = CSng(mValue) / mMax
    L = (View.Width - 9) * P + 4
    T = (View.Height - 3) / 2
    
    If View.Enabled Then
        If bHover Then
            If bDown Then
                Pen1.Color = ThemeColor.AccentDarkColor
            Else
                Pen1.Color = ThemeColor.AccentLightColor
            End If
        Else
            Pen1.Color = ThemeColor.AccentColor
        End If
        Pen2.Color = ThemeColor.BKColor
    Else
        Pen1.Color = ThemeColor.BKDrakColor
        Pen2.Color = ThemeColor.BKColor
    End If
    Brush.Color = Pen1.Color
    Pen1.Width = 2
    Pen2.Width = 2
    
    Canvas.SetSmoothingMode SmoothingModeNone
    Canvas.DrawLine Pen2, 4, T, View.Width - 5, T
    Canvas.DrawLine Pen1, 4, T, L, T
    
    Canvas.SetSmoothingMode SmoothingModeAntiAlias
    If bHover Then
        Canvas.FillEllipse Brush.GetBaseBrush, L - 4.5, T - 4.5, 9, 9
    Else
        Canvas.FillEllipse Brush.GetBaseBrush, L - 3.5, T - 3.5, 7, 7
    End If
End Sub

