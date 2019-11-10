VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Bitmap          As cBitmap
Private ViewList        As New cObjectList
Private HoverView       As cView
Private FocusView       As cView
Private CaptureView     As cView
Private Graphics        As cGraphics
Private Activity        As cActivity

Private mBKColor        As Long
Private FirstDone       As Boolean

Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseWheel(ByVal Forward As Boolean)
Public Event MouseHover()
Public Event MouseLeave()
Public Event KeyDown(ByVal KeyCode As Integer)
Public Event KeyUp(ByVal KeyCode As Integer)
Public Event Resize()

Private Sub Class_Terminate(): Release: End Sub

Friend Sub Create(Root As cActivity, ByVal nWidth As Long, ByVal nHeight As Long)
    Set Bitmap = New cBitmap
    Bitmap.CreateEmptyBitmap nWidth, nHeight
    Set Graphics = Bitmap.GetGraphics
    Set Activity = Root
End Sub

Friend Sub Release()
    Dim i As Long
    Dim View As cView
    
    Set Graphics = Nothing
    Set Bitmap = Nothing
    Set FocusView = Nothing
    Set HoverView = Nothing
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            View.Release
        Set View = Nothing
    Next
    Set ViewList = Nothing
End Sub

Friend Property Get IsFirstDone() As Boolean: IsFirstDone = FirstDone: End Property
Public Property Get ViewCount() As Long: ViewCount = ViewList.GetCount: End Property

Public Property Get BackgroundColor() As Long: BackgroundColor = mBKColor: End Property
Public Property Let BackgroundColor(ByVal nColor As Long): mBKColor = nColor: Refresh: End Property

Public Function GetActivity() As cActivity: Set GetActivity = Activity: End Function
Friend Function GetRoot() As cActivity: Set GetRoot = Activity: End Function
Friend Sub SetRoot(Root As cActivity)
    If Root Is Nothing Then Exit Sub
    Set Activity = Root
End Sub

Friend Function GetImage() As cImage: Set GetImage = Bitmap.ToImage: End Function
Friend Sub ChangeSize(ByVal nWidth As Long, ByVal nHeight As Long)
    If nWidth <> Bitmap.Width Or nHeight <> Bitmap.Height Then
        Set Graphics = Nothing
        Set Bitmap = Nothing
        Set Bitmap = New cBitmap
        Bitmap.CreateEmptyBitmap nWidth, nHeight
        Set Graphics = Bitmap.GetGraphics
        RaiseEvent Resize
    End If
    Refresh
End Sub

Friend Function HandleMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef Result As Long) As Boolean
    If uMsg = WM_MOUSEMOVE Or uMsg = WM_LBUTTONDOWN Or uMsg = WM_RBUTTONDOWN Or uMsg = WM_MBUTTONDOWN Or uMsg = WM_MOUSEWHEEL Then
        OnMouseEvent uMsg, wParam, lParam
    ElseIf uMsg = WM_LBUTTONUP Then
        wParam = MK_LBUTTON
        OnMouseEvent uMsg, wParam, lParam
    ElseIf uMsg = WM_MBUTTONUP Then
        wParam = MK_MBUTTON
        OnMouseEvent uMsg, wParam, lParam
    ElseIf uMsg = WM_RBUTTONUP Then
        wParam = MK_RBUTTON
        OnMouseEvent uMsg, wParam, lParam
    ElseIf uMsg = WM_MOUSEHOVER Then
        RaiseEvent MouseHover
    ElseIf uMsg = WM_MOUSELEAVE Then
        SetHoverView Nothing
        SetCaptureView Nothing
        RaiseEvent MouseLeave
    ElseIf uMsg = WM_KEYDOWN Or uMsg = WM_KEYUP Then
        OnKeyEvent uMsg, wParam, lParam
    ElseIf uMsg = WM_CHAR Then
        OnCharEvent uMsg, wParam, lParam
    Else
        HandleMessage = False
        Exit Function
    End If
    
    HandleMessage = True
End Function

Friend Sub SetHoverView(View As cView)
    If Not View Is HoverView Then
        If Not HoverView Is Nothing Then
            HoverView.RaiseMouseLeave
        End If
        Set HoverView = View
        If Not HoverView Is Nothing Then HoverView.RaiseMouseHover
    End If
End Sub

Friend Sub SetFocusView(View As cView)
    If Not View Is FocusView Then
        If Not FocusView Is Nothing Then
            FocusView.RaiseLostFocus
        End If
        Set FocusView = View
        If Not FocusView Is Nothing Then FocusView.RaiseGotFocus
    End If
End Sub

Friend Sub SetCaptureView(View As cView)
    Set CaptureView = View
End Sub

Public Sub ReDrawAllView()
    FirstDone = False
    Refresh
End Sub

Public Sub Refresh()
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
    
    Graphics.Clear mBKColor
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                If Not FirstDone Then View.ReDraw
                View.MergeImage
                CM.m(3, 3) = View.Opacity
                Attr.SetColorMatrix CM, ColorAdjustTypeBitmap
                Graphics.DrawImageRectRect View.GetImage, _
                                        NewRectF(View.Left, View.Top, View.Width, View.Height), _
                                        NewRectF(0, 0, View.Width, View.Height), Attr

            End If
        Set View = Nothing
    Next
    
    If ViewList.GetCount > 0 Then FirstDone = True
    
    If Not Activity Is Nothing Then Activity.Refresh
End Sub

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

    Graphics.SetClip L, T, W, H, CombineModeReplace
    Graphics.Clear mBKColor
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.CheckInRect(L, T, W, H) Then
                View.RefreshRect L - View.Left, T - View.Top, W, H
                CM.m(3, 3) = View.Opacity
                Attr.SetColorMatrix CM, ColorAdjustTypeBitmap
                Graphics.DrawImageRectRect View.GetImage, _
                                        NewRectF(View.Left, View.Top, View.Width, View.Height), _
                                        NewRectF(0, 0, View.Width, View.Height), Attr
            End If
        Set View = Nothing
    Next
    
    Graphics.ResetClip
    If Not Activity Is Nothing Then Activity.RefreshRect L, T, W, H
End Sub

Private Sub OnMouseEvent(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Dim Button As Integer
    Dim Shift As Integer
    Dim X As Integer, Y As Integer
    Dim i As Long
    Dim View As cView
    
    If uMsg <> WM_MOUSEWHEEL Then
        Button = GetButton(wParam)
        Shift = GetShift(wParam)
        X = LowWord(lParam)
        Y = HighWord(lParam)
    End If
    
    If uMsg = WM_MOUSEMOVE Then
        If Not CaptureView Is Nothing Then
            If CaptureView.Visible And CaptureView.Enabled Then
                Set View = CaptureView.GetParent
                Do While Not View Is Nothing
                    X = X - View.Left
                    Y = Y - View.Top
                    Set View = View.GetParent
                Loop
                CaptureView.RaiseMouseMove Button, Shift, X - CaptureView.Left, Y - CaptureView.Top
                Exit Sub
            End If
        End If
    
        For i = ViewList.GetCount - 1 To 0 Step -1
            Set View = ViewList.GetObject(i)
                If View.Visible Then
                    If View.CheckMouseIn(X, Y) Then
                        If View.Enabled Then View.RaiseMouseMove Button, Shift, X - View.Left, Y - View.Top
                        Set View = Nothing
                        Exit Sub
                    End If
                End If
            Set View = Nothing
        Next
        If Not HoverView Is Nothing Then
            If HoverView.Enabled Then HoverView.RaiseMouseLeave
            Set HoverView = Nothing
        End If
        RaiseEvent MouseMove(Button, Shift, X, Y)
    ElseIf uMsg = WM_LBUTTONDOWN Or uMsg = WM_RBUTTONDOWN Or uMsg = WM_MBUTTONDOWN Then
        For i = ViewList.GetCount - 1 To 0 Step -1
            Set View = ViewList.GetObject(i)
                If View.Visible Then
                    If View.CheckMouseIn(X, Y) Then
                        If View.Enabled Then View.RaiseMouseDown Button, Shift, X - View.Left, Y - View.Top
                        Set View = Nothing
                        Exit Sub
                    End If
                End If
            Set View = Nothing
        Next
        If Not FocusView Is Nothing Then
            If FocusView.Enabled Then FocusView.RaiseLostFocus
            Set FocusView = Nothing
        End If
        SetHoverView Nothing
        SetCaptureView Nothing
        RaiseEvent MouseDown(Button, Shift, X, Y)
    ElseIf uMsg = WM_LBUTTONUP Or uMsg = WM_RBUTTONUP Or uMsg = WM_MBUTTONUP Then
        If Not FocusView Is Nothing Then
            If FocusView.Visible And FocusView.Enabled Then
                Set View = FocusView.GetParent
                Do While Not View Is Nothing
                    X = X - View.Left
                    Y = Y - View.Top
                    Set View = View.GetParent
                Loop
                FocusView.RaiseMouseUp Button, Shift, X, Y
            End If
        Else
            RaiseEvent MouseUp(Button, Shift, X, Y)
        End If
    ElseIf uMsg = WM_MOUSEWHEEL Then
        If Not HoverView Is Nothing Then
            If HoverView.Enabled Then HoverView.RaiseMouseWheel IIf(wParam > 0, True, False)
        Else
            RaiseEvent MouseWheel(IIf(wParam > 0, True, False))
        End If
    End If
End Sub

Public Sub OnKeyEvent(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    If Not FocusView Is Nothing Then
        If FocusView.Enabled Then
            If uMsg = WM_KEYDOWN Then FocusView.RaiseKeyDown wParam
            If uMsg = WM_KEYUP Then FocusView.RaiseKeyUp wParam
        End If
    Else
        If uMsg = WM_KEYDOWN Then RaiseEvent KeyDown(wParam)
        If uMsg = WM_KEYUP Then RaiseEvent KeyUp(wParam)
    End If
End Sub

Public Sub OnCharEvent(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Dim Ascii As Long
    Static LastChar As Long
    
    If wParam <= 127 Then
        LastChar = 0
    Else
        If LastChar = 0 Then
        LastChar = wParam
        Exit Sub
        End If
    End If
    
    Ascii = CLng(LastChar) * 256 + wParam
    LastChar = 0
    
    If Not FocusView Is Nothing Then
        If FocusView.Enabled Then FocusView.RaiseChar Ascii
    End If
End Sub

Public Function CreateView(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Dim View As New cView
    View.Create Me, Nothing, L, T, W, H
    ViewList.Add View
    Set CreateView = View
End Function

Public Function GetView(ByVal Index As Long): Set GetView = ViewList.GetObject(Index): End Function

Public Sub RemoveView(View As cView): f_RemoveView View: Refresh: End Sub
Friend Sub f_RemoveView(View As cView)
    If ViewList.Find(View) = -1 Then Exit Sub
    View.SetRoot Nothing
    View.f_SetParent Nothing
    ViewList.RemoveObject View:
End Sub

Public Sub Clear(): ViewList.Clear: End Sub

Public Sub AddView(View As cView)
    Dim Root As cLayout
    Dim Parent As cView
    
    Set Root = View.GetRoot
    Set Parent = View.GetParent
    
    If Not Root Is Nothing Then
        If Not Root Is Me Then
            Root.f_RemoveView View
            Root.Refresh
            View.SetRoot Me
            ViewList.Add View
        End If
    Else
        View.SetRoot Me
        ViewList.Add View
    End If
    
    If Not Parent Is Nothing Then
        Parent.f_RemoveView View
    End If
    View.f_SetParent Nothing
    
    View.ReDraw
    Me.Refresh
End Sub

Friend Sub RelayMsgToTopView(ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Dim i As Long
    Dim View As cView
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            View.RaiseHandleMsg nMsg, wParam, lParam
        Set View = Nothing
    Next
End Sub
