VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ClassStyle      As Long
Private WindowStyle     As Long
Private WindowStyleEx   As Long
Private WindowIcon      As Long
Private hInstance       As Long
Private mHwnd           As Long
Private mCaption        As String
Private mLeft           As Long
Private mTop            As Long
Private mWidth          As Long
Private mHeight         As Long
Private mBorderStyle    As BorderStyleEnum
Private IsHasMin        As Boolean
Private isHasMax        As Boolean

Private Const ClassName       As String = "RFWindow"

Private Sub Class_Initialize()
    ClassStyle = CS_OWNDC Or CS_HREDRAW Or CS_VREDRAW Or CS_DBLCLKS
    Icon = LoadIconA(ByVal 0&, IDI_APPLICATION)
    Caption = "Window"
End Sub

Public Property Get CloseButton() As Boolean
    CloseButton = ((ClassStyle And CS_NOCLOSE) = CS_NOCLOSE)
End Property

Public Property Let CloseButton(ByVal IsHas As Boolean)
    If IsHas Then
        ClassStyle = ClassStyle Or CS_NOCLOSE
    Else
        ClassStyle = ClassStyle Or CS_NOCLOSE Xor CS_NOCLOSE
    End If
End Property

Public Property Get MinButton() As Boolean: MinButton = IsHasMin: End Property
Public Property Let MinButton(ByVal IsHas As Boolean): IsHasMin = IsHas: End Property
Public Property Get MaxButton() As Boolean: MaxButton = isHasMax: End Property
Public Property Let MaxButton(ByVal IsHas As Boolean): isHasMax = IsHas: End Property

Public Property Get Icon() As Long: Icon = WindowIcon: End Property
Public Property Let Icon(ByVal nHandle As Long): WindowIcon = nHandle: End Property

Public Property Get AppInstance() As Long: AppInstance = hInstance: End Property
Public Property Let AppInstance(ByVal nInstance As Long): hInstance = nInstance: End Property

Public Property Get BorderStyle() As BorderStyleEnum: BorderStyle = mBorderStyle: End Property
Public Property Let BorderStyle(ByVal nStyle As BorderStyleEnum): mBorderStyle = nStyle: End Property

Public Property Get hwnd() As Long
    hwnd = mHwnd
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal sCaption As String)
    mCaption = sCaption
    If hwnd <> 0 Then
        SetWindowTextA mHwnd, sCaption
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = IsWindowEnabled(mHwnd)
End Property

Public Property Let Enabled(ByVal bEnable As Boolean)
    EnableWindow mHwnd, IIf(bEnable, 1, 0)
End Property

Public Property Get Left() As Long: Left = mLeft: End Property
Public Property Let Left(ByVal nLeft As Long)
    mLeft = nLeft
    Move mLeft, mTop, Width, Height
End Property

Public Property Get Top() As Long: Top = mTop: End Property
Public Property Let Top(ByVal nTop As Long)
    mTop = nTop
    Move mLeft, mTop, Width, Height
End Property

Public Property Get Width() As Long: Width = mWidth: End Property
Public Property Let Width(ByVal nWidth As Long)
    mWidth = nWidth
    Move mLeft, mTop, Width, Height
End Property

Public Property Get Height() As Long: Height = mHeight: End Property
Public Property Let Height(ByVal nHeight As Long)
    mHeight = nHeight
    Move mLeft, mTop, Width, Height
End Property

Public Property Get HDC() As Long
    HDC = GetDC(mHwnd)
End Property

Public Sub Move(ByVal L As Long, Optional T, Optional W, Optional H)
    mLeft = L
    If Not IsMissing(T) Then mTop = T
    If Not IsMissing(W) Then mWidth = W
    If Not IsMissing(H) Then mHeight = H
    
    SetWindowPos mHwnd, HWND_NOTOPMOST, L, T, W, H, SWP_NOZORDER
End Sub

Public Function Build() As cWindow
    If Not RegisterWindowClass Then GoTo Fail
    If Not CreateWindow Then GoTo Fail
    Set Build = Me
Fail:
End Function

Private Function RegisterWindowClass() As Boolean
    Dim WC As WNDCLASSA
    
    With WC
        .Style = ClassStyle
        .hbrBackground = COLOR_WINDOW
        .hCursor = LoadCursorA(0&, IDC_ARROW)
        .hIcon = WindowIcon
        .lpszClassName = ClassName
        .hInstance = hInstance
        .lpfnWndProc = GetAddress(AddressOf WndProc)
    End With
    
    RegisterWindowClass = RegisterClassA(WC) <> 0
End Function

Private Function CreateWindow() As Boolean
    Select Case mBorderStyle
        Case BorderStyleEnum.None
            WindowStyle = WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED
            WindowStyleEx = WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR
        Case BorderStyleEnum.FixedSingle
            WindowStyle = WS_CAPTION Or WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED Or WS_SYSMENU
            If IsHasMin Then WindowStyle = WindowStyle Or WS_MINIMIZEBOX
            If isHasMax Then WindowStyle = WindowStyle Or WS_MAXIMIZEBOX
            WindowStyleEx = WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR Or WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
        Case BorderStyleEnum.Sizable
            WindowStyle = WS_CAPTION Or WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED Or WS_SYSMENU Or WS_THICKFRAME
            If IsHasMin Then WindowStyle = WindowStyle Or WS_MINIMIZEBOX
            If isHasMax Then WindowStyle = WindowStyle Or WS_MAXIMIZEBOX
            WindowStyleEx = WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR Or WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
        Case BorderStyleEnum.FixedDialog
            WindowStyle = WS_CAPTION Or WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED Or WS_SYSMENU Or &H80
            WindowStyleEx = WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR Or WS_EX_WINDOWEDGE Or WS_EX_DLGMODALFRAME
        Case BorderStyleEnum.FixedToolWindow
            WindowStyle = WS_CAPTION Or WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED Or WS_SYSMENU
            WindowStyleEx = WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR Or WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW
        Case BorderStyleEnum.SizableToolWindow
            WindowStyle = WS_CAPTION Or WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED Or WS_SYSMENU Or WS_THICKFRAME
            WindowStyleEx = WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR Or WS_EX_WINDOWEDGE Or WS_EX_DLGMODALFRAME
    End Select

    mHwnd = CreateWindowExA(WindowStyleEx, ClassName, mCaption, WindowStyle, mLeft, mTop, mWidth, mHeight, ByVal 0, ByVal 0, hInstance, ByVal 0)
    If mHwnd = 0 Then GoTo Fail
    CreateWindow = True
    Exit Function
Fail:
End Function

Friend Function HandleMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

End Function
