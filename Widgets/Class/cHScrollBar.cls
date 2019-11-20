VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHScrollBar"
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
Public Event Scrolled(ByVal Offset As Long)

Private bDown           As Boolean
Private bHover          As Boolean

Private mValue          As Long
Private mMax            As Long

Private ThemeColor  As ThemeColorType

Private Enum HoverFlag
    HF_None
    HF_UP
    HF_MID
    HF_DOWN
End Enum

Private Const MinWidth  As Single = 36
Private Const MinHeigth As Single = 18

Private mHF         As HoverFlag
Private BarSize     As Single
Private BarLeft      As Single
Private LastX       As Integer

Private Sub Class_Initialize()
    mMax = 10
    SetTheme Preset
End Sub

Public Sub Create(Parent As Object, ByVal sName As String, ByVal nLeft As Integer, ByVal nTop As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer)
    nWidth = IIf(nWidth < MinWidth, MinWidth, nWidth)
    nHeight = IIf(nHeight < MinHeigth, MinHeigth, nHeight)
    Set View = Parent.CreateView(nLeft, nTop, nWidth, nHeight)
    View.Name = sName
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
    '-----------------------------------------
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
    nWidth = IIf(nWidth < MinWidth, MinWidth, nWidth)
    If View.Width = nWidth Then Exit Property
    View.Width = nWidth
End Property

Public Property Get Height() As Integer: Height = View.Height: End Property
Public Property Let Height(ByVal nHeight As Integer)
    nHeight = IIf(nHeight < MinHeigth, MinHeigth, nHeight)
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
    End Select
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

Public Property Get Value() As Long: Value = mValue: End Property
Public Property Let Value(ByVal nValue As Long)
    nValue = IIf(nValue < 0, 0, IIf(nValue > mMax, mMax, nValue))
    If mValue = nValue Then Exit Property
    RaiseEvent Scrolled(nValue - mValue)
    mValue = nValue
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
    mHF = HF_None
    bHover = False
    View.Refresh
End Sub

Private Sub View_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim Offset As Integer
    RaiseEvent MouseDown(Button, Shift, X, Y)
    bDown = True

    If X < 18 Then
        mHF = HF_UP
    ElseIf X > View.Width - 19 Then
        mHF = HF_DOWN
    ElseIf X > BarLeft And X <= BarLeft + BarSize Then
        mHF = HF_MID
    Else
        mHF = HF_None
    End If
        
    Select Case mHF
        Case HoverFlag.HF_UP
            If mValue > 0 Then
                mValue = mValue - 1
                RaiseEvent Scrolled(-1)
                View.Refresh
            End If
        Case HoverFlag.HF_MID
            LastX = X - BarLeft
            View.Refresh
        Case HoverFlag.HF_DOWN
            If mValue < mMax Then
                mValue = mValue + 1
                RaiseEvent Scrolled(1)
                View.Refresh
            End If
        Case HoverFlag.HF_None
            If View.Width <= 54 Then Exit Sub
            Offset = mMax * 0.1
            If Offset = 0 Then Offset = 1
            
            If X < BarLeft Then
                If mValue > 0 Then
                    mValue = mValue - Offset
                    If mValue < 0 Then mValue = 0
                    RaiseEvent Scrolled(-Offset)
                    View.Refresh
                End If
            ElseIf X > BarLeft + BarSize Then
                If mValue < mMax Then
                    mValue = mValue + Offset
                    If mValue > mMax Then mValue = mMax
                    RaiseEvent Scrolled(Offset)
                    View.Refresh
                End If
            End If
    End Select
End Sub

Private Sub View_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim nHF As HoverFlag
    Dim OffsetX As Single
    Dim nValue As Long
    Dim n As Long
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If bDown Then
        If mHF = HF_MID Then
            OffsetX = X - 18 - LastX
            n = View.Width - 36 - BarSize
            n = IIf(n <= 0, 1, n)
            nValue = (OffsetX / n) * mMax
            
            nValue = IIf(nValue < 0, 0, IIf(nValue > mMax, mMax, nValue))
            If mValue <> nValue Then
                RaiseEvent Scrolled(nValue - mValue)
                mValue = nValue
                View.Refresh
            End If
        End If
    Else
        If X < 18 Then
            nHF = HF_UP
        ElseIf X > View.Width - 19 Then
            nHF = HF_DOWN
        ElseIf X > BarLeft And X <= BarLeft + BarSize Then
            nHF = HF_MID
        Else
            nHF = HF_None
        End If
        
        If mHF <> nHF Then
            mHF = nHF
            View.Refresh
        End If
    End If
End Sub

Private Sub View_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    bDown = False
    View.Refresh
End Sub

Private Sub View_MouseWheel(ByVal Forward As Boolean)
    Dim Offset As Integer
    Offset = mMax * 0.05
    If Offset = 0 Then Offset = 1
    mHF = HF_None
    
    If Forward Then
        Me.Value = Me.Value - Offset
    Else
        Me.Value = Me.Value + Offset
    End If
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Brush As New cSolidBrush
    Dim Brush2 As New cSolidBrush
    Dim Path As New cPath
    Dim OffsetT As Single
    
    If View.Enabled Then
        If bHover Then
            Brush.Color = ThemeColor.BorderLightColor
        Else
            Brush.Color = ThemeColor.BorderColor
        End If
    Else
        Brush.Color = ThemeColor.BorderDisEnableColor
    End If
    
    If bDown Then
        Brush2.Color = ThemeColor.AccentDarkColor
    Else
        Brush2.Color = ThemeColor.AccentColor
    End If
    
    OffsetT = (View.Height - 18) / 2
    BarSize = (View.Width - 36) / mMax
    BarSize = IIf(BarSize < 15, 15, BarSize)
    BarLeft = 18 + (View.Width - 36 - BarSize) * (mValue / mMax)
    
    If bHover Then
        Canvas.Clear ThemeColor.BKDrakColor
    End If
    
    '左三角
    Path.AddLine 12, OffsetT + 2, 6, OffsetT + 8
    Path.AddLine 6, OffsetT + 8, 12, OffsetT + 14
    Path.CloseFigure
    
    If mHF = HF_UP Then
        Canvas.fillPath Brush2.GetBaseBrush, Path
    Else
        Canvas.fillPath Brush.GetBaseBrush, Path
    End If
    
    Path.Reset
    '右三角
    Path.AddLine View.Width - 18 + 5, OffsetT + 2, View.Width - 18 + 11, OffsetT + 8
    Path.AddLine View.Width - 18 + 11, OffsetT + 8, View.Width - 18 + 5, OffsetT + 14
    Path.CloseFigure
    
    If mHF = HF_DOWN Then
        Canvas.fillPath Brush2.GetBaseBrush, Path
    Else
        Canvas.fillPath Brush.GetBaseBrush, Path
    End If
    
    '控制条
    If View.Width > 54 Then
        If mHF = HF_MID Then
            Canvas.FillRectangle Brush2.GetBaseBrush, BarLeft, OffsetT + 3, BarSize, 11
        Else
            Canvas.FillRectangle Brush.GetBaseBrush, BarLeft, OffsetT + 3, BarSize, 11
        End If
    Else
        BarSize = 0
    End If
End Sub

