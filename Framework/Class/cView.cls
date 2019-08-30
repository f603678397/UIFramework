VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Bitmap          As cBitmap
Private ResultBMP       As cBitmap
Private Graphics        As cGraphics                '自身画布
Private Canvas          As cGraphics                '合并图画布
Private mName           As String
Private mTop            As Integer
Private mLeft           As Integer
Private mWidth          As Integer
Private mHeight         As Integer
Private mParent         As cView
Private mVisible        As Boolean
Private ViewList        As New cObjectList
Private Layout          As cLayout
Private isHover         As Boolean
Private isFocus         As Boolean
Private mBKColor        As Long

Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseWheel(ByVal Forward As Boolean)
Public Event MouseHover()
Public Event MouseLeave()
Public Event GotFocus()
Public Event LostFocus()
Public Event Paint(Canvas As cGraphics)

'Private Sub Class_Terminate(): Set Bitmap = Nothing: End Sub

Public Property Get Left() As Integer: Left = mLeft: End Property
Public Property Let Left(ByVal nLeft As Integer): mLeft = nLeft: Notify: End Property

Public Property Get Top() As Integer: Top = mTop: End Property
Public Property Let Top(ByVal nTop As Integer): mTop = nTop: Notify: End Property

Public Property Get Width() As Integer: Width = mWidth: End Property
Public Property Let Width(ByVal nWidth As Integer): mWidth = nWidth: RaisePaint: End Property

Public Property Get Height() As Integer: Height = mHeight: End Property
Public Property Let Height(ByVal nHeight As Integer): mHeight = nHeight: RaisePaint: End Property

Public Property Get Visible() As Boolean: Visible = mVisible: End Property
Public Property Let Visible(ByVal bVisible As Boolean): mVisible = bVisible: RaisePaint: End Property

Public Property Get Name() As String: Name = mName: End Property
Public Property Let Name(ByVal sName As String): mName = sName: End Property

Public Property Get BackgroundColor() As Long: BackgroundColor = mBKColor: End Property
Public Property Let BackgroundColor(ByVal nColor As Long): mBKColor = nColor: RaisePaint: End Property

Friend Sub Create(Root As cLayout, ParentView As cView, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Set Bitmap = New cBitmap
    mLeft = L: mTop = T
    mWidth = W: mHeight = H
    Bitmap.CreateEmptyBitmap W, H
    Set Graphics = Bitmap.GetGraphics
    ResultBMP.CreateEmptyBitmap W, H
    Set Canvas = ResultBMP.GetGraphics
    
    mVisible = True
    Set Layout = Root
    Set mParent = ParentView
    RaisePaint
End Sub

Friend Function GetImage() As cImage: Set GetImage = ResultBMP.ToImage: End Function

Friend Function GetRoot() As cLayout: Set GetRoot = Layout: End Function
'Friend Sub SetRoot(Root As cLayout): Set Layout = Root: End Sub

Friend Sub RaiseLostFocus(): isFocus = False: RaiseEvent LostFocus: End Sub
Friend Sub RaiseGotFocus()
    If Not isFocus Then
        isFocus = True
        RaiseEvent GotFocus
    End If
End Sub

Friend Sub RaiseMouseLeave(): isHover = False: RaiseEvent MouseLeave: End Sub
Friend Sub RaiseMouseHover()
    If Not isHover Then
        isHover = True
        RaiseEvent MouseHover
    End If
End Sub

Friend Sub RaiseMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim i As Long
    Dim View As cView
    For i = ViewList.GetCount - 1 To 0 Step -1
        Set View = ViewList.GetObject(i)
            If View.CheckMouseIn(X, Y) Then
                View.RaiseMouseMove Button, Shift, X - View.Left, Y - View.Top
                Set View = Nothing
                Exit Sub
            End If
        Set View = Nothing
    Next
    Layout.SetHoverView Me
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim i As Long
    Dim View As cView
    For i = ViewList.GetCount - 1 To 0 Step -1
        Set View = ViewList.GetObject(i)
            If View.CheckMouseIn(X, Y) Then
                View.RaiseMouseDown Button, Shift, X - View.Left, Y - View.Top
                Set View = Nothing
                Exit Sub
            End If
        Set View = Nothing
    Next
    Layout.SetFocusView Me
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseWheel(ByVal Forward As Boolean)
    RaiseEvent MouseWheel(Forward)
End Sub

Friend Sub RaisePaint()
    Graphics.Clear mBKColor
    RaiseEvent Paint(Graphics)
    MergeImage
    Notify
End Sub

Private Sub Notify()
    If Layout.IsFirstDone Then
        If Not mParent Is Nothing Then
            mParent.NotifyRefesh mLeft, mTop, mWidth, mHeight
        Else
            Layout.RefreshRect mLeft, mTop, mWidth, mHeight
        End If
    End If
End Sub

Friend Function CheckMouseIn(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X >= mLeft And X <= mLeft + mWidth And Y >= mTop And Y <= mTop + mHeight Then CheckMouseIn = True
End Function

Friend Sub NotifyRefesh(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    If Not mParent Is Nothing Then
        mParent.NotifyRefesh L + mLeft, T + mTop, W, H
    Else
        Layout.RefreshRect L, T, W, H
    End If
End Sub

Friend Function CheckInRect(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As Boolean
    Dim W1 As Long, H1 As Long
    Dim W2 As Long, H2 As Long
    
    W1 = mWidth + W
    H1 = mHeight + H
    
    If L <= mLeft Then
        W2 = mLeft + mWidth - L
    Else
        W2 = L + W - mLeft
    End If
    
    If T <= mTop Then
        H2 = mTop + mHeight - T
    Else
        H2 = T + H - mTop
    End If
    
    If W2 < W1 And H2 < H1 Then CheckInRect = True
End Function

Friend Sub RefreshRect(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Dim i As Long
    Dim View As cView
    
    Canvas.SetClip L, T, W, H, CombineModeReplace
    Canvas.Clear
    Canvas.DrawImage Bitmap.ToImage, 0, 0
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                If View.CheckInRect(L, T, W, H) Then
                    View.RefreshRect L - View.Left, T - View.Top, W, H
                    Canvas.DrawImage View.GetImage, View.Left, View.Top
                End If
            End If
        Set View = Nothing
    Next
    Canvas.ResetClip
End Sub

Friend Sub MergeImage()
    Dim i As Long
    Dim View As cView
    
    Canvas.Clear
    Canvas.DrawImage Bitmap.ToImage, 0, 0
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                View.MergeImage
                Canvas.DrawImage View.GetImage, View.Left, View.Top
            End If
        Set View = Nothing
    Next
End Sub

Public Function GetParent() As cView: Set GetParent = mParent: End Function
'Public Sub SetParent(NewParent As cView)

'End Sub

Public Sub Move(ByVal L As Integer, Optional ByVal T, Optional ByVal W, Optional ByVal H)
    Dim bReDraw As Boolean
    mLeft = L
    mTop = IIf(IsMissing(T), mTop, T)
    If Not IsMissing(W) And mWidth <> W Then: mWidth = W: bReDraw = True
    If Not IsMissing(H) And mHeight <> H Then mHeight = H: bReDraw = True
    
    If bReDraw Then RaisePaint
End Sub

Public Function CreateView(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Dim View As New cView
    View.Create Layout, Me, L, T, W, H
    ViewList.Add View
    Set CreateView = View
End Function

Public Sub RemoveView(View As cView): ViewList.RemoveObject View: End Sub

Public Sub Clear(): ViewList.Clear: End Sub
