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
Private ReDrawRect      As RECTI
Private FirstDraw       As Boolean

Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseWheel(ByVal Forward As Boolean)
Public Event MouseHover()
Public Event MouseLeave()
Public Event Click()
Public Event DblClick()
Public Event GotFocus()
Public Event LostFocus()
Public Event Paint(Canvas As cGraphics)

Private Sub Class_Terminate(): Release: End Sub

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

Public Property Get Visible() As Boolean: Visible = mVisible: End Property
Public Property Let Visible(ByVal bVisible As Boolean): mVisible = bVisible: Notify: End Property

Public Property Get Name() As String: Name = mName: End Property
Public Property Let Name(ByVal sName As String): mName = sName: End Property

Public Property Get BackgroundColor() As Long: BackgroundColor = mBKColor: End Property
Public Property Let BackgroundColor(ByVal nColor As Long): mBKColor = nColor: ReDraw: MergeImageSelf: Notify: End Property

Public Property Get ZOrder() As Long: ZOrder = mParent.GetChildZOrder(Me): End Property
Public Property Let ZOrder(ByVal nOrder As Long)
    mParent.ChangeChildZOrder Me, nOrder
    Notify
End Property

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
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            View.Release
        Set View = Nothing
    Next
    
    Set ViewList = Nothing
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
    RaiseEvent MouseMove(Button, Shift, X - mLeft, Y - mTop)
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
    RaiseEvent MouseDown(Button, Shift, X - mLeft, Y - mTop)
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

Friend Sub RaiseMouseWheel(ByVal Forward As Boolean)
    RaiseEvent MouseWheel(Forward)
End Sub

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
    If Not mParent Is Nothing Then
        mParent.NotifyRefesh L + mLeft, T + mTop, W, H
    Else
        Layout.RefreshRect L + mLeft, T + mTop, W, H
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
    
    If Not FirstDraw Then ReDraw
    
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

Friend Sub ReDraw()
    Graphics.Clear mBKColor
    RaiseEvent Paint(Graphics)
    FirstDraw = True
End Sub

Friend Sub MergeImageSelf()
    Dim i As Long
    Dim View As cView
    
    Canvas.Clear
    Canvas.DrawImage Bitmap.ToImage, 0, 0
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                Canvas.DrawImage View.GetImage, View.Left, View.Top
            End If
        Set View = Nothing
    Next
End Sub

Friend Sub MergeImage()
    Dim i As Long
    Dim View As cView
    
    Canvas.Clear
    Canvas.DrawImage Bitmap.ToImage, 0, 0
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                If Not Layout.IsFirstDone Then View.ReDraw
                View.MergeImage
                Canvas.DrawImage View.GetImage, View.Left, View.Top
            End If
        Set View = Nothing
    Next
End Sub

Public Function GetParent() As cView: Set GetParent = mParent: End Function
'Public Sub SetParent(NewParent As cView)

'End Sub

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

Public Sub RemoveView(View As cView): ViewList.RemoveObject View: End Sub

Public Sub Clear(): ViewList.Clear: End Sub
