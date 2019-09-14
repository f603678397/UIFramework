VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Graphics        As cGraphics
Private PreFunc         As Long
Private mhWnd           As Long
Private mWidth          As Long
Private mHeight         As Long
Private SrcDC           As Long
Private DstDC           As Long
Private mBMP            As Long
Private mbLayer         As Boolean

Private srcPoint            As POINTAPI
Private BlendFunc32bpp      As BLENDFUNCTION
Private winSize             As Size

Private bLeaveTrack         As Boolean
Private bHoverTrack         As Boolean

Private Layout              As cLayout
Private bCreated            As Boolean

Private Sub Class_Terminate(): Release: End Sub

Public Property Get hWnd() As Long: hWnd = mhWnd: End Property

Public Property Get Width() As Long: Width = mWidth: End Property
Public Property Get Height() As Long: Height = mHeight: End Property

Public Function CreateLayout() As cLayout
    Dim mLayout As New cLayout
    mLayout.Create Me, mWidth, mHeight
    Set CreateLayout = mLayout
End Function

Public Function GetLayout() As cLayout: Set GetLayout = Layout: End Function
Public Sub SetLayout(NewLayout As cLayout)
    If NewLayout Is Nothing Then Exit Sub
    If NewLayout Is Layout Then Exit Sub
    If Not Layout Is Nothing Then
        Layout.SetRoot Nothing
    End If
    Set Layout = NewLayout
    Layout.SetRoot Me
    Refresh
End Sub

Public Sub Refresh()
    Graphics.Clear cColor.Transparent
    If Not Layout Is Nothing Then Graphics.DrawImage Layout.GetImage, 0, 0
    If mbLayer Then
        UpdateLayeredWindow mhWnd, SrcDC, ByVal 0&, winSize, DstDC, srcPoint, 0, BlendFunc32bpp, ULW_ALPHA
    Else
        BitBlt SrcDC, 0, 0, mWidth, mHeight, DstDC, 0, 0, vbSrcCopy
    End If
End Sub

Friend Sub RefreshRect(ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer)
    Dim RECT As Drawing2D.RECTF

    If Not Layout Is Nothing Then
        RECT.Left = L: RECT.Top = T: RECT.Right = W: RECT.bottom = H
        Graphics.SetClip L, T, W, H, CombineModeReplace
        Graphics.Clear cColor.Transparent
        Graphics.DrawImagePointRect Layout.GetImage, L, T, RECT
        Graphics.ResetClip
    Else
        Exit Sub
    End If
    
    If mbLayer Then
        UpdateLayeredWindow mhWnd, SrcDC, ByVal 0&, winSize, DstDC, srcPoint, 0, BlendFunc32bpp, ULW_ALPHA
    Else
        BitBlt SrcDC, 0, 0, mWidth, mHeight, DstDC, 0, 0, vbSrcCopy
    End If
End Sub

Friend Sub Create(ByVal hWnd As Long, ByVal bLayer As Boolean)
    mbLayer = bLayer
    mhWnd = hWnd
    PreFunc = SetWindowLongA(hWnd, GWL_WNDPROC, AddressOf WndProc)
    GetWindowSize mhWnd
    SrcDC = GetDC(mhWnd)
    CreateMemoryDC
    Set Graphics = Drawing.CreateGraphicsFromHWND(DstDC)
    bHoverTrack = True
    bCreated = True
End Sub

Friend Sub Release()
    If Not bCreated Then Exit Sub
    SetWindowLongA mhWnd, GWL_WNDPROC, PreFunc
    Graphics.Release
    If Not Layout Is Nothing Then Layout.Release
    Set Graphics = Nothing
    Set Layout = Nothing
    DestroyMemoryDC
    ActivityList.RemoveObject Me
    bCreated = False
End Sub

Friend Function HandleMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Result As Long
    Select Case uMsg
        Case WM_PAINT
            onPaint
        Case WM_SIZE
            onSize wParam, lParam
        Case WM_MOUSEMOVE
            onMouseMove wParam, lParam
        Case WM_MOUSEHOVER
            onMouseHover wParam, lParam
        Case WM_MOUSELEAVE
            onMouseLeave
        Case Else
            If Not Layout Is Nothing Then
                If Layout.HandleMessage(uMsg, wParam, lParam, Result) Then
                    HandleMessage = Result
                Else
                    HandleMessage = CallWindowProcA(PreFunc, mhWnd, uMsg, wParam, lParam)
                End If
            Else
                HandleMessage = CallWindowProcA(PreFunc, mhWnd, uMsg, wParam, lParam)
            End If
    End Select
End Function

Private Sub GetWindowSize(ByVal hWnd As Long)
    Dim mRect As RECT
    GetClientRect hWnd, mRect
    mWidth = mRect.Right - mRect.Left
    mHeight = mRect.bottom - mRect.Top
    
    srcPoint.X = 0
    srcPoint.Y = 0
    winSize.cx = mWidth
    winSize.cy = mHeight
    
    With BlendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
End Sub

Private Sub CreateMemoryDC()
    Dim bmpInfo As BITMAPINFO
    With bmpInfo.bmiHeader
        .biBitCount = 32
        .biWidth = mWidth
        .biHeight = mHeight
        .biSize = LenB(bmpInfo.bmiHeader)
        .biSizeImage = mWidth * mHeight * (.biBitCount / 8)
        .biPlanes = 1
    End With
    
    DstDC = CreateCompatibleDC(SrcDC)
    mBMP = CreateDIBSection(DstDC, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
    SelectObject DstDC, mBMP
End Sub

Private Sub DestroyMemoryDC()
    DeleteObject mBMP
    DeleteObject DstDC
End Sub

Private Sub onPaint()
    Dim PS As PAINTSTRUCT
    BeginPaint mhWnd, PS
    Refresh
    EndPaint mhWnd, PS
End Sub

Private Sub onSize(ByVal wParam As Long, ByVal lParam As Long)
    If wParam = SIZE_RESTORED Or SIZE_MAXIMIZED Then
        DestroyMemoryDC
        
        mWidth = LowWord(lParam)
        mHeight = HighWord(lParam)

        winSize.cx = mWidth
        winSize.cy = mHeight
    
        SrcDC = GetDC(mhWnd)
        CreateMemoryDC
        Graphics.Release
        Set Graphics = Nothing
        Set Graphics = cDrawing.CreateGraphicsFromHDC(DstDC)
        If Not Layout Is Nothing Then Layout.ChangeSize mWidth, mHeight
    End If
End Sub

Private Sub onMouseMove(ByVal wParam As Long, ByVal lParam As Long)
    Dim TME As TRACKMOUSEEVENT
    Dim Flags As Long

    If bLeaveTrack Then
        Flags = TME_LEAVE
    End If

    If bHoverTrack Then
        Flags = Flags Or TME_HOVER
    End If

    With TME
        .cbSize = LenB(TME)
        .hWndTrack = mhWnd
        .dwFlags = Flags
        .dwHoverTime = HOVER_DEFAULT
    End With

    Call TRACKMOUSEEVENT(TME)
    
    If Not Layout Is Nothing Then
        Layout.HandleMessage WM_MOUSEMOVE, wParam, lParam, 0
    End If
End Sub

Private Sub onMouseHover(ByVal wParam As Long, ByVal lParam As Long)
    bHoverTrack = False
    bLeaveTrack = True
    If Not Layout Is Nothing Then
        Layout.HandleMessage WM_MOUSEHOVER, 0, 0, 0
    End If
End Sub

Private Sub onMouseLeave()
    bLeaveTrack = False
    bHoverTrack = True
    If Not Layout Is Nothing Then
        Layout.HandleMessage WM_MOUSELEAVE, 0, 0, 0
    End If
End Sub
