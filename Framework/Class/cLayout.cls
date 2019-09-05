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
Private Graphics        As cGraphics
Private Activity        As cActivity

Private mBKColor        As Long
Private FirstDone       As Boolean
Private bForeground     As Boolean

Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseWheel(ByVal Forward As Boolean)
Public Event MouseHover()
Public Event MouseLeave()

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
    
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            View.Release
        Set View = Nothing
    Next
    Set ViewList = Nothing
End Sub

Friend Property Get IsFirstDone() As Boolean: IsFirstDone = FirstDone: End Property
Public Property Get ViewCount(): ViewCount = ViewList.GetCount: End Property

Public Property Get IsForeground() As Boolean: IsForeground = bForeground: End Property
Friend Property Let IsForeground(ByVal B As Boolean): bForeground = B: End Property

Public Property Get BackgroundColor() As Long: BackgroundColor = mBKColor: End Property
Public Property Let BackgroundColor(ByVal nColor As Long): mBKColor = nColor: Refresh: End Property

Friend Function GetImage() As cImage: Set GetImage = Bitmap.ToImage: End Function
Friend Sub ChangeSize(ByVal nWidth As Long, ByVal nHeight As Long)
    Set Graphics = Nothing
    Set Bitmap = Nothing
    Set Bitmap = New cBitmap
    Bitmap.CreateEmptyBitmap nWidth, nHeight
    Set Graphics = Bitmap.GetGraphics
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
        RaiseEvent MouseLeave
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
        HoverView.RaiseMouseHover
    End If
End Sub

Friend Sub SetFocusView(View As cView)
    If Not View Is FocusView Then
        If Not FocusView Is Nothing Then
            FocusView.RaiseLostFocus
        End If
        Set FocusView = View
        FocusView.RaiseGotFocus
    End If
End Sub

Public Sub Refresh()
    Dim i As Long
    Dim View As cView
    
    Graphics.Clear mBKColor
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.Visible Then
                If Not FirstDone Then View.ReDraw
                View.MergeImage
                Graphics.DrawImage View.GetImage, View.Left, View.Top
            End If
        Set View = Nothing
    Next
    
    If ViewList.GetCount > 0 Then FirstDone = True
    
    If bForeground Then Activity.Refresh
End Sub

Friend Sub RefreshRect(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Dim i As Long
    Dim View As cView

    Graphics.SetClip L, T, W, H, CombineModeReplace
    Graphics.Clear mBKColor
    For i = 0 To ViewList.GetCount - 1
        Set View = ViewList.GetObject(i)
            If View.CheckInRect(L, T, W, H) Then
                View.RefreshRect L - View.Left, T - View.Top, W, H
                Graphics.DrawImage View.GetImage, View.Left, View.Top
            End If
        Set View = Nothing
    Next
    
    If bShowRedrawRgn Then
        Dim Brush As New cSolidBrush
        Brush.Color = cColor.FromARGB(100, 255, 0, 0)
        Graphics.FillRectangle Brush.GetBaseBrush, 0, 0, Bitmap.Width, Bitmap.Height
    End If
    
    Graphics.ResetClip
    If bForeground Then Activity.RefreshRect L, T, W, H
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
        For i = ViewList.GetCount - 1 To 0 Step -1
            Set View = ViewList.GetObject(i)
                If View.Visible Then
                    If View.CheckMouseIn(X, Y) Then
                        View.RaiseMouseMove Button, Shift, X, Y
                        Set View = Nothing
                        Exit Sub
                    End If
                End If
            Set View = Nothing
        Next
        If Not HoverView Is Nothing Then
            HoverView.RaiseMouseLeave
            Set HoverView = Nothing
        End If
        RaiseEvent MouseMove(Button, Shift, X, Y)
    ElseIf uMsg = WM_LBUTTONDOWN Or uMsg = WM_RBUTTONDOWN Or uMsg = WM_MBUTTONDOWN Then
        For i = ViewList.GetCount - 1 To 0 Step -1
            Set View = ViewList.GetObject(i)
                If View.Visible Then
                    If View.CheckMouseIn(X, Y) Then
                        View.RaiseMouseDown Button, Shift, X - View.Left, Y - View.Top
                        Set View = Nothing
                        Exit Sub
                    End If
                End If
            Set View = Nothing
        Next
        If Not FocusView Is Nothing Then
            FocusView.RaiseLostFocus
            Set FocusView = Nothing
        End If
        RaiseEvent MouseDown(Button, Shift, X, Y)
    ElseIf uMsg = WM_LBUTTONUP Or uMsg = WM_RBUTTONUP Or uMsg = WM_MBUTTONUP Then
        If Not FocusView Is Nothing Then
            If FocusView.Visible Then
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
            HoverView.RaiseMouseWheel IIf(wParam > 0, True, False)
        Else
            RaiseEvent MouseWheel(IIf(wParam > 0, True, False))
        End If
    ElseIf uMsg = WM_MOUSEHOVER Then
        RaiseEvent MouseHover
    ElseIf uMsg = WM_MOUSELEAVE Then
        RaiseEvent MouseLeave
    End If
End Sub

Public Function CreateView(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Dim View As New cView
    View.Create Me, Nothing, L, T, W, H
    ViewList.Add View
    Set CreateView = View
End Function

Public Sub RemoveView(View As cView): ViewList.RemoveObject View: End Sub

Public Sub Clear(): ViewList.Clear: End Sub

'Public Sub AddView(View As cView, ByVal L As Integer, ByVal T As Integer)
'
'End Sub
