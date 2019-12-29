VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents View          As cView
Attribute View.VB_VarHelpID = -1
Private Bitmap                   As New cBitmap

Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseHover()
Public Event MouseLeave()
Public Event GotFocus()
Public Event LostFocus()
Public Event Click()
Public Event DblClick()
Public Event ItemClick(ByVal ItemIndex As Long)

Private Adapter         As cListAdapter
Private ItemBitmap      As New cObjectList
Private ThemeColor      As ThemeColorType

Private ItemHeight      As Integer
Private OffsetY         As Single
Private bDown           As Boolean
Private LastIndex       As Long

Private Sub Class_Initialize()
    ItemHeight = 15
    LastIndex = -1
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
    LastIndex = -1
    
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
        Case LCase("BackColor")
            Me.BackColor = CLng(Value)
        Case LCase("BorderColor")
            Me.BorderColor = CLng(Value)
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

Public Sub Refresh(): View.Refresh: End Sub

Public Sub SetTheme(NewTheme As PresetThemeEnum)
    Static bNotFirst As Boolean
    
    Select Case NewTheme
        Case PresetThemeEnum.DrakTheme
            SetDarkTheme ThemeColor
        Case PresetThemeEnum.LightTheme
            SetLightTheme ThemeColor
    End Select
    
    If bNotFirst Then RefreshAllItem
    bNotFirst = True
End Sub

'=========================================================================================
Public Property Get BorderColor() As Long: BorderColor = ThemeColor.BorderColor: End Property
Public Property Let BorderColor(ByVal nColor As Long)
    If nColor = ThemeColor.BorderColor Then Exit Property
    
    ThemeColor.BorderColor = nColor
    
    View.Refresh
End Property

Public Function GetItemHeight() As Long: GetItemHeight = ItemHeight: End Function
Public Sub SetItemHeight(ByVal nHeight As Integer)
    If ItemHeight = nHeight Then Exit Sub
    
    ItemHeight = nHeight
    
    RefreshAllItem
End Sub

Public Sub SetAdapter(ListAdapter As cListAdapter)
    Set Adapter = ListAdapter
    
    ItemBitmap.Clear
    CreateItemBitmap
    MergeBitmap
End Sub

Private Sub RefreshAllItem()
    If Adapter Is Nothing Then Exit Sub
    If Adapter.GetCount = 0 Then Exit Sub
    
    ItemBitmap.Clear
    CreateItemBitmap
    MergeBitmap
End Sub

Private Sub RefreshItem(ByVal nIndex As Long, Optional ByVal State As ItemState = ItemState.IS_NORMAL)
    Dim Bmp As cBitmap
    Dim Canvas As cGraphics
    
    If nIndex < 0 Or nIndex > ItemBitmap.GetCount - 1 Then Exit Sub
    
    Set Bmp = ItemBitmap.GetObject(nIndex)
    Set Canvas = Bmp.GetGraphics
    Canvas.Clear
    Adapter.DrawItem Canvas, nIndex, State
    MergeBitmap
    
    View.Refresh
End Sub

Private Sub CreateItemBitmap()
    Dim Bmp As cBitmap
    Dim i As Long
    
    For i = 0 To Adapter.GetCount - 1
        Set Bmp = New cBitmap
        Bmp.CreateEmptyBitmap View.Width - 4, ItemHeight
        Adapter.DrawItem Bmp.GetGraphics, i, IS_NORMAL
        ItemBitmap.Add Bmp
    Next
End Sub

Private Sub MergeBitmap()
    Dim BmpHeight As Long
    Dim Bmp As cBitmap
    
    Dim Canvas As cGraphics
    Dim T As Long
    
    Dim i As Long
    
    For i = 0 To ItemBitmap.GetCount - 1
        Set Bmp = ItemBitmap.GetObject(i)
        BmpHeight = BmpHeight + Bmp.Height
    Next

    Bitmap.CreateEmptyBitmap View.Width - 4, BmpHeight
    Set Canvas = Bitmap.GetGraphics

    For i = 0 To ItemBitmap.GetCount - 1
        Set Bmp = ItemBitmap.GetObject(i)
        Canvas.DrawImage Bmp.ToImage, 0, T
        T = T + Bmp.Height
    Next
End Sub

Private Function GetItemIndex(ByVal Y As Integer) As Long
    Dim T As Long
    Dim i As Long
    
    Y = Y - 2 + Abs(OffsetY)
    
    For i = 0 To ItemBitmap.GetCount - 1
        If Y >= T And Y < T + ItemHeight Then
            GetItemIndex = i
            Exit Function
        End If
        T = T + ItemHeight
    Next
    GetItemIndex = -1
End Function

Private Sub View_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim ItemIndex As Long
    
    ItemIndex = GetItemIndex(Y)
    
    If ItemIndex > -1 Then
        RefreshItem ItemIndex, IS_DOWN
        LastIndex = ItemIndex
        bDown = True
    End If
End Sub

Private Sub View_MouseLeave()
    If LastIndex > -1 Then
        RefreshItem LastIndex, IS_NORMAL
        LastIndex = -1
    End If
End Sub

Private Sub View_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim ItemIndex As Long
    
    If Not bDown Then
        ItemIndex = GetItemIndex(Y)
        If ItemIndex > -1 And ItemIndex <> LastIndex Then
            RefreshItem LastIndex, IS_NORMAL
            RefreshItem ItemIndex, IS_HOVER
            LastIndex = ItemIndex
        End If
    End If
End Sub

Private Sub View_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    If bDown Then
        RefreshItem LastIndex, IS_HOVER
        RaiseEvent ItemClick(LastIndex)
        bDown = False
    End If
End Sub

Private Sub View_MouseWheel(ByVal Forward As Boolean)
    If Bitmap.Height < View.Height - 4 Then Exit Sub
    
    If Forward Then
        If OffsetY < 0 Then
            OffsetY = OffsetY + Int(ItemHeight / 2)
            If OffsetY > 0 Then OffsetY = 0
        Else
            Exit Sub
        End If
    Else
        If (View.Height - 4) - OffsetY < Bitmap.Height Then
            OffsetY = OffsetY - Int(ItemHeight / 2)
            If (View.Height - 4) - OffsetY > Bitmap.Height Then
                OffsetY = View.Height - 4 - Bitmap.Height
            End If
        Else
            Exit Sub
        End If
    End If
    
    View.Refresh
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Dim Pen     As New cPen
    Dim Brush   As New cSolidBrush
    Dim IA      As New cImageAttributes
    Dim CM      As ColorMatrix
    
    Pen.Color = ThemeColor.BorderColor
    Canvas.Clear ThemeColor.BKColor
    
    CM.M(0, 0) = 1: CM.M(1, 1) = 1: CM.M(2, 2) = 1
    If View.Enabled Then
        CM.M(3, 3) = 1
    Else
        CM.M(3, 3) = 0.5
    End If
    IA.SetColorMatrix CM
    
    Canvas.DrawRectangle Pen, 0, 0, View.Width - 1, View.Height - 1
    Canvas.SetClip 2, 2, View.Width - 4, View.Height - 4, CombineModeReplace
    Canvas.DrawImageRectRect Bitmap.ToImage, _
                            NewRectF(2, 2 + OffsetY, Bitmap.Width, Bitmap.Height), _
                            NewRectF(0, 0, Bitmap.Width, Bitmap.Height), _
                            IA
    Canvas.ResetClip
End Sub
