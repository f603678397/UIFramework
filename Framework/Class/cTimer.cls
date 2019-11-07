VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mhWnd           As Long
Private mHandle         As Long
Private mID             As Long
Private mEnabled        As Boolean
Private mInterval       As Long

Public Event onTime()

Private Sub Class_Initialize(): mEnabled = True: End Sub
Private Sub Class_Terminate(): Release: End Sub

Public Function Create(ByVal hWnd As Long) As Boolean
    mhWnd = hWnd
    mID = ObjPtr(Me)
    
    mHandle = SetTimer(mhWnd, mID, 1000, AddressOf TimerProc)
    If mHandle <> 0 Then
        TimerList.Add Me
        Create = True
    Else
        Create = False
    End If
    KillTimer mhWnd, mHandle
End Function

Public Sub Release()
    If mHandle = 0 Then Exit Sub
    KillTimer mhWnd, mHandle
    TimerList.Remove TimerList.Find(Me)
    mHandle = 0
End Sub

Friend Sub CallOnTime(): RaiseEvent onTime: End Sub

Friend Property Get ID() As Long: ID = mID: End Property

Public Property Get Enabled() As Boolean: Enabled = mEnabled: End Property
Public Property Let Enabled(ByVal bEnabled As Boolean)
    If bEnabled = mEnabled Then Exit Property
    
    mEnabled = bEnabled
    If mEnabled Then
        If mInterval > 0 Then
            mHandle = SetTimer(mhWnd, mID, mInterval, AddressOf TimerProc)
        End If
    Else
        KillTimer mhWnd, mHandle
    End If
End Property

Public Property Get Interval() As Long: Interval = mInterval: End Property
Public Property Let Interval(ByVal n As Long)
    If n = mInterval Then Exit Property
    
    mInterval = n
    If mInterval > 0 Then
        If mEnabled Then
            mHandle = SetTimer(mhWnd, mID, mInterval, AddressOf TimerProc)
        End If
    Else
        KillTimer mhWnd, mHandle
    End If
End Property
