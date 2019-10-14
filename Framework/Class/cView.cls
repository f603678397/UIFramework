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
Private mOpacity        As Single
Private mParent         As cView
Private mVisible        As Boolean
Private mEnabled        As Boolean
Private mPaddingLeft    As Integer
Private mPaddingTop     As Integer
Private mPaddingRight   As Integer
Private mPaddingBottom  As Integer
Private ViewList        As New cObjectList
Private Layout          As cLayout
Private isHover         As Boolean
Private mIsFocus         As Boolean
Private mBKColor        As Long
Private ReDrawRect      As RECTI
Private FirstDraw       As Boolean
Private mTag            As String

Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseWheel(ByVal Forward As Boolean)
Public Event MouseHover()
Public Event MouseLeave()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(ByVal KeyCode As Integer)
Public Event KeyUp(ByVal KeyCode As Integer)
Public Event KeyPress(ByVal KeyAscii As Long)
Public Event GotFocus()
Public Event LostFocus()
Public Event Paint(Canvas As cGraphics)
Public Event HandleMsg(ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Private Sub Class_Terminate(): Release: End Sub

Public Property Get Tag() As String: Tag = mTag: End Property
Public Property Let Tag(ByVal sTag As String): mTag = sTag: End Property

Public Property Get Left() As Integer: Left = mLeft: End Property
Public Property Let Left(ByVal nLeft As Integer)
    If nLeft = mLeft Then Exit Property
    With ReDrawRect
        .Left = IIf(mLeft < nLeft, mLeft, nLeft)
        .Top = mTop
        .Width = mWidth + Abs(nLeft - mLeft)
        .Height = mHeight
    End With
    mLeft = nLeft
    Notify True
End Property

Public Property Get Top() As Integer: Top = mTop: End Property
Public Property Let Top(ByVal nTop As Integer)
    If nTop = mTop Then Exit Property
    With ReDrawRect
        .Left = mLeft
        .Top = IIf(mTop < nTop, mTop, nTop)
        .Width = mWidth
        .Height = mHeight + Abs(nTop - mTop)
    End With
    mTop = nTop
    Notify True
End Property

Public Property Get Width() As Integer: Width = mWidth: End Property
Public Property Let Width(ByVal nWidth As Integer)
    If nWidth = mWidth Then Exit Property
    With ReDrawRect
        .Left = mLeft
        .Top = mTop
        .Width = IIf(nWidth < mWidth, mWidth, nWidth)
        .Height = mHeight
    End With
    mWidth = nWidth
    ChangeSize
    ReDraw
    MergeImageSelf
    Notify True
End Property

Public Property Get Height() As Integer: Height = mHeight: End Property
Public Property Let Height(ByVal nHeight As Integer)
    If nHeight = mHeight Then Exit Property
    With ReDrawRect
        .Left = mLeft
        .Top = mTop
        .Width = mWidth
        .Height = IIf(nHeight < mHeight, mHeight, nHeight)
    End With
    mHeight = nHeight
    ChangeSize
    ReDraw
    MergeImageSelf
    Notify True
End Property

Public Property Let Padding(ByVal nPadding As Integer)
    mPaddingLeft = nPadding
    mPaddingTop = nPadding
    mPaddingRight = nPadding
    mPaddingBottom = nPadding
    Notify
End Property

Public Property Get PaddingLeft() As Integer: PaddingLeft = mPaddingLeft: End Property
Public Property Let PaddingLeft(ByVal nPadding As Integer): mPaddingLeft = nPadding: Notify: End Property

Public Property Get PaddingTop() As Integer: PaddingTop = mPaddingTop: End Property
Public Property Let PaddingTop(ByVal nPadding As Integer): mPaddingTop = nPadding: Notify: End Property

Public Property Get PaddingRight() As Integer: PaddingRight = mPaddingRight: End Property
Public Property Let PaddingRight(ByVal nPadding As Integer): mPaddingRight = nPadding: Notify: End Property

Public Property Get PaddingBottom() As Integer: PaddingBottom = mPaddingBottom: End Property
Public Property Let PaddingBottom(ByVal nPadding As Integer): mPaddingBottom = nPadding: Notify: End Property

Public Property Get IsFocus(): IsFocus = mIsFocus: End Property

Public Property Get Enabled() As Boolean: Enabled = mEnabled: End Property
Public Property Let Enabled(ByVal bEnabled As Boolean): mEnabled = bEnabled: End Property

Public Property Get Visible() As Boolean: Visible = mVisible: End Property
Public Property Let Visible(ByVal bVisible As Boolean): mVisible = bVisible: Notify: End Property

Public Property Get Name() As String: Name = mName: End Property
Public Property Let Name(ByVal sName As String): mName = sName: End Property

Public Property Get BackgroundColor() As Long: BackgroundColor = mBKColor: End Property
Public Property Let BackgroundColor(ByVal nColor As Long): mBKColor = nColor: ReDraw: MergeImageSelf: Notify: End Property

Public Property Get ZOrder() As Long: ZOrder = mParent.GetChildZOrder(Me): End Property
Public Property Let ZOrder(ByVal nOrder As Long): mParent.ChangeChildZOrder Me, nOrder: Notify: End Property

Public Property Get Opacity() As Single: Opacity = mOpacity: End Property
Public Property Let Opacity(ByVal fOpacity As Single): mOpacity = fOpacity: Notify: End Property

Public Property Get ViewCount() As Long: ViewCount = ViewList.GetCount: End Property

Friend Sub Create(Root As cLayout, ParentView As cView, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Set Bitmap = New cBitmap
    Set ResultBMP = New cBitmap
    mLeft = L: mTop = T
    mWidth = W: mHeight = H
    Bitmap.CreateEmptyBitmap W, H
    Set Graphics = Bitmap.GetGraphics
    ResultBMP.CreateEmptyBitmap W, H
    Set Canvas = ResultBMP.GetGraphics
    
    mVisible = True
    mOpacity = 1
    mEnabled = True
    Set Layout = Root
    Set mParent = ParentView
End Sub

Friend Function GetChildZOrder(View As cView) As Long: GetChildZOrder = ViewList.Find(View): End Function
Friend Sub ChangeChildZOrder(View As cView, ByVal Position As Long): ViewList.MoveObject View, Position: End Sub

Private Sub ChangeSize()
    Set Graphics = Nothing
    Set Canvas = Nothing
    
    Bitmap.CreateEmptyBitmap mWidth, mHeight
    Set Graphics = Bitmap.GetGraphics
    ResultBMP.CreateEmptyBitmap mWidth, mHeight
    Set Canvas = ResultBMP.GetGraphics
End Sub

Friend Sub Release()
    Dim i As Long
    Dim View As cView
    
    Set Graphics = Nothing
    Set Canvas = Nothing
    Set Bitmap = Nothing
    Set ResultBMP = Nothing
    Set Layout = Nothing
    Set mParent = Nothing
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            View.Release
        Set View = Nothing
    Next
    
    Set ViewList = Nothing
End Sub

Friend Function GetImage() As cImage: Set GetImage = ResultBMP.ToImage: End Function

Friend Function GetRoot() As cLayout: Set GetRoot = Layout: End Function
Friend Sub SetRoot(Root As cLayout): Set Layout = Root: End Sub

Friend Sub RaiseLostFocus(): mIsFocus = False: RaiseEvent LostFocus: End Sub
Friend Sub RaiseGotFocus()
    If Not mIsFocus Then
        mIsFocus = True
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
                If View.Enabled Then View.RaiseMouseMove Button, Shift, X - View.Left, Y - View.Top
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
                If View.Enabled Then View.RaiseMouseDown Button, Shift, X - View.Left, Y - View.Top
                Set View = Nothing
                Exit Sub
            End If
        Set View = Nothing
    Next
    Layout.SetFocusView Me
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim mX As Integer, mY As Integer
    Dim CurrTime As Long
    Static LastTime As Long
    
    mX = X - mLeft: mY = Y - mTop
    RaiseEvent MouseUp(Button, Shift, mX, mY)
    
    If mX >= 0 And mX <= mWidth And mY >= 0 And mY <= mHeight Then
        CurrTime = GetTickCount
        If CurrTime - LastTime < GetDoubleClickTime Then
            RaiseEvent DblClick
            LastTime = 0
        Else
            RaiseEvent Click
            LastTime = CurrTime
        End If
    End If
End Sub

Friend Sub RaiseMouseWheel(ByVal Forward As Boolean): RaiseEvent MouseWheel(Forward): End Sub

Friend Sub RaiseKeyUp(ByVal KeyCode As Integer): RaiseEvent KeyUp(KeyCode): End Sub
Friend Sub RaiseKeyDown(ByVal KeyCode As Integer): RaiseEvent KeyDown(KeyCode): End Sub

Friend Sub RaiseChar(ByVal Ascii As Long): RaiseEvent KeyPress(Ascii): End Sub

Friend Sub RaisePaint()
    ReDraw
    MergeImageSelf
    Notify
End Sub

Private Sub Notify(Optional ByVal NewRgn As Boolean = False)
    If Layout.IsFirstDone Then
        If Not mParent Is Nothing Then
            If Not NewRgn Then
                mParent.NotifyRefesh mLeft, mTop, mWidth, mHeight
            Else
                mParent.NotifyRefesh ReDrawRect.Left, ReDrawRect.Top, ReDrawRect.Width, ReDrawRect.Height
            End If
        Else
            If Not NewRgn Then
                Layout.RefreshRect mLeft, mTop, mWidth, mHeight
            Else
                Layout.RefreshRect ReDrawRect.Left, ReDrawRect.Top, ReDrawRect.Width, ReDrawRect.Height
            End If
        End If
    End If
End Sub

Friend Function CheckMouseIn(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X >= mLeft And X <= mLeft + mWidth And Y >= mTop And Y <= mTop + mHeight Then CheckMouseIn = True
End Function

Friend Sub NotifyRefesh(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Dim mL As Integer, mT As Integer, mW As Integer, mH As Integer
    Dim cW As Integer, cH As Integer
    
    If L > mWidth - mPaddingRight Or T > mHeight - mPaddingBottom Or W < 1 Or H < 1 Then Exit Sub
    
    cW = mWidth - mPaddingLeft - mPaddingRight
    cH = mHeight - mPaddingTop - mPaddingBottom
    
    mL = IIf(L < mPaddingLeft, mPaddingLeft, L)
    mT = IIf(T < mPaddingTop, mPaddingTop, T)
    mW = IIf(W > cW, cW, W)
    mH = IIf(H > cH, cH, H)
    
    If L >= mPaddingLeft And L + W >= mWidth - mPaddingRight Then mW = cW - (L - mPaddingLeft)
    If T >= mPaddingTop And T + H >= mHeight - mPaddingBottom Then mH = cH - (T - mPaddingTop)
    
    If L <= mPaddingLeft And L + W < mWidth - mPaddingRight Then mW = W - (mPaddingLeft - L)
    If T <= mPaddingTop And T + H < mHeight - mPaddingBottom Then mH = H - (mPaddingTop - T)
    
    If Not mParent Is Nothing Then
        mParent.NotifyRefesh mL + mLeft, mT + mTop, mW, mH
    Else
        Layout.RefreshRect mL + mLeft, mT + mTop, mW, mH
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
    Dim Attr As New cImageAttributes
    Dim CM As Drawing2D.ColorMatrix
    
    With CM
        .m(0, 0) = 1
        .m(1, 1) = 1
        .m(2, 2) = 1
        .m(3, 3) = 1
        .m(4, 4) = 1
    End With
    
    If Not FirstDraw Then ReDraw
    
    Canvas.SetClip mPaddingLeft, mPaddingTop, mWidth - mPaddingLeft - PaddingRight, mHeight - mPaddingTop - mPaddingBottom, CombineModeReplace
    Canvas.SetClip L, T, W, H, CombineModeIntersect
    Canvas.Clear
    Canvas.DrawImage Bitmap.ToImage, 0, 0
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                If View.CheckInRect(L, T, W, H) Then
                    View.RefreshRect L - View.Left, T - View.Top, W, H
                    CM.m(3, 3) = View.Opacity
                    Attr.SetColorMatrix CM, ColorAdjustTypeBitmap
                    Canvas.DrawImageRectRect View.GetImage, _
                                        NewRectF(View.Left, View.Top, View.Width, View.Height), _
                                        NewRectF(0, 0, View.Width, View.Height), Attr
                End If
            End If
        Set View = Nothing
    Next
    Canvas.ResetClip
End Sub

Friend Sub ReDraw()
    Graphics.Clear mBKColor
    RaiseEvent Paint(Graphics)
    FirstDraw = True
End Sub

Friend Sub MergeImageSelf()
    Dim i As Long
    Dim View As cView
    Dim Attr As New cImageAttributes
    Dim CM As Drawing2D.ColorMatrix
    
    With CM
        .m(0, 0) = 1
        .m(1, 1) = 1
        .m(2, 2) = 1
        .m(3, 3) = 1
        .m(4, 4) = 1
    End With
    
    Canvas.Clear
    Canvas.DrawImage Bitmap.ToImage, 0, 0
    Canvas.SetClip mPaddingLeft, mPaddingTop, mWidth - mPaddingLeft - PaddingRight, mHeight - mPaddingTop - mPaddingBottom, CombineModeReplace
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                CM.m(3, 3) = View.Opacity
                Attr.SetColorMatrix CM, ColorAdjustTypeBitmap
                Canvas.DrawImageRectRect View.GetImage, _
                                    NewRectF(View.Left, View.Top, View.Width, View.Height), _
                                    NewRectF(0, 0, View.Width, View.Height), Attr
            End If
        Set View = Nothing
    Next
    
    Canvas.ResetClip
End Sub

Friend Sub MergeImage()
    Dim i As Long
    Dim View As cView
    Dim Attr As New cImageAttributes
    Dim CM As Drawing2D.ColorMatrix
    
    With CM
        .m(0, 0) = 1
        .m(1, 1) = 1
        .m(2, 2) = 1
        .m(3, 3) = 1
        .m(4, 4) = 1
    End With
    
    Canvas.Clear
    Canvas.DrawImage Bitmap.ToImage, 0, 0
    Canvas.SetClip mPaddingLeft, mPaddingTop, mWidth - mPaddingLeft - PaddingRight, mHeight - mPaddingTop - mPaddingBottom, CombineModeReplace
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                If Not Layout.IsFirstDone Then View.ReDraw
                View.MergeImage
                CM.m(3, 3) = View.Opacity
                Attr.SetColorMatrix CM, ColorAdjustTypeBitmap
                Canvas.DrawImageRectRect View.GetImage, _
                                    NewRectF(View.Left, View.Top, View.Width, View.Height), _
                                    NewRectF(0, 0, View.Width, View.Height), Attr
            End If
        Set View = Nothing
    Next
    
    Canvas.ResetClip
End Sub

Public Function GetCacheBitmap() As cBitmap: Set GetCacheBitmap = ResultBMP: End Function
Public Function GetCacheBitmapWithoutChild() As cBitmap: Set GetCacheBitmapWithoutChild = Bitmap: End Function

Public Function GetParent() As cView: Set GetParent = mParent: End Function
Public Sub SetParent(NewParent As cView)
    If NewParent Is mParent Then Exit Sub
    
    If NewParent Is Nothing Then
        If Not Layout Is Nothing Then
            Layout.AddView Me
        End If
    Else
        NewParent.AddView Me
    End If
End Sub

Friend Sub f_SetParent(NewParent As cView): Set mParent = NewParent: End Sub

Public Sub Refresh()
    RaisePaint
End Sub

Public Sub Move(ByVal L As Integer, Optional ByVal T, Optional ByVal W, Optional ByVal H)
    Dim mL As Integer, mT As Integer, mW As Integer, mH As Integer
    Dim bReDraw As Boolean, bMove As Boolean
    
    mL = L
    mT = IIf(IsMissing(T), mTop, T)
    mW = IIf(IsMissing(W), mWidth, W)
    mH = IIf(IsMissing(H), mHeight, H)
    
    If mL <> mLeft Or mT <> mTop Then bMove = True
    If mW <> mWidth Or mH <> mHeight Then bReDraw = True
    
    With ReDrawRect
        .Left = IIf(mL < mLeft, mL, mLeft)
        .Top = IIf(mT < mTop, mT, mTop)
        .Width = IIf(mLeft + mWidth > mL + mW, mLeft + mWidth - .Left, mL + mW - .Left)
        .Height = IIf(mTop + mHeight > mT + mH, mTop + mHeight - .Top, mT + mH - .Top)
    End With
    
    mLeft = mL: mTop = mT: mWidth = mW: mHeight = mH
    
    If bReDraw Then ChangeSize: ReDraw: MergeImageSelf
    Notify bMove
End Sub

Public Function CreateView(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Dim View As New cView
    View.Create Layout, Me, L, T, W, H
    ViewList.Add View
    Set CreateView = View
    View.ReDraw
End Function

Public Sub AddView(View As cView)
    Dim mRoot As cLayout
    Dim Parent As cView
    
    Set mRoot = View.GetRoot
    Set Parent = View.GetParent
    
    If Parent Is Me Then Exit Sub
    If Parent Is Nothing Then
        If Not mRoot Is Nothing Then
            mRoot.f_RemoveView View
            If Not mRoot Is Layout Then mRoot.Refresh
        End If
    Else
        Parent.f_RemoveView View
    End If
    View.f_SetParent Me
    ViewList.Add View
    
    View.SetRoot Layout
    Layout.Refresh
End Sub

Public Function GetView(ByVal Index As Long): Set GetView = ViewList.GetObject(Index): End Function

Public Sub RemoveView(View As cView): f_RemoveView View: Notify: End Sub
Friend Sub f_RemoveView(View As cView)
    If ViewList.Find(View) = -1 Then Exit Sub
    View.SetRoot Nothing
    View.f_SetParent Nothing
    ViewList.RemoveObject View
End Sub

Public Sub Clear(): ViewList.Clear: End Sub

Friend Sub RelayMsgToTopView(ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Dim i As Long
    Dim View As cView
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            View.RaiseHandleMsg nMsg, wParam, lParam
        Set View = Nothing
    Next
End Sub

Friend Sub RaiseHandleMsg(ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    RaiseEvent HandleMsg(nMsg, wParam, lParam)
End Sub

Public Sub SendMsgToSameLevel(ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    If Not mParent Is Nothing Then
        mParent.RelayMsgToTopView nMsg, wParam, lParam
    Else
        Layout.RelayMsgToTopView nMsg, wParam, lParam
    End If
End Sub
