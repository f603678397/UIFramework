VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Initialize()
    Drawing.Init
End Sub

Public Sub Terminate()
    Dim i As Long
    Dim Activity    As cActivity
    Dim Timer       As cTimer
    Dim Drawable    As cDrawable

    For i = 0 To ActivityList.GetCount - 1
        Set Activity = ActivityList.GetObject(i)
            Activity.Release
        Set Activity = Nothing
    Next
    Set ActivityList = Nothing

    For i = 0 To TimerList.GetCount - 1
        Set Timer = TimerList.GetObject(i)
            If Not Timer Is Nothing Then Timer.Release
        Set Timer = Nothing
    Next
    Set TimerList = Nothing

    For i = 0 To DrawableList.GetCount - 1
        Set Drawable = DrawableList.GetObject(i)
            Drawable.Release
        Set Drawable = Nothing
    Next
    Set DrawableList = Nothing

    IM.DisposeAll
    Set IM = Nothing
    Drawing.SHUTDOWN
End Sub

Public Function CreateActivity(ByVal hWnd As Long, Optional ByVal bLayer As Boolean = False) As cActivity
    Dim Activity As cActivity
    Dim i As Long
    
    For i = 0 To ActivityList.GetCount - 1
        Set Activity = ActivityList.GetObject(i)
            If Activity.hWnd = hWnd Then
                Set CreateActivity = Activity
                Exit Function
            End If
        Set Activity = Nothing
    Next
    
    Set Activity = New cActivity
    Activity.Create hWnd, bLayer
    ActivityList.Add Activity
    Set CreateActivity = Activity
End Function

Public Sub DestroyActivity(Activity As cActivity)
    Activity.Release
End Sub

Public Function GetImageManager() As cImageManager: Set GetImageManager = IM: End Function

Public Sub ShowRedrawRgn(ByVal bShow As Boolean)
    bShowRedrawRgn = bShow
End Sub

Public Sub ShowViewEdge(ByVal bShow As Boolean)
    bShowViewEdge = bShow
End Sub
