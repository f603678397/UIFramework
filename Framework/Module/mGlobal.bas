Attribute VB_Name = "mGlobal"
Option Explicit

Public Drawing As New cDrawing
Public TimerList As New cObjectList
Public ActivityList As New cObjectList

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Activity As cActivity
    Dim i As Long
    
    For i = 0 To ActivityList.GetCount - 1
        Set Activity = ActivityList.GetObject(i)
            If Activity.hWnd = hWnd Then
                WndProc = Activity.HandleMessage(uMsg, wParam, lParam)
                Set Activity = Nothing
                Exit Function
            End If
        Set Activity = Nothing
    Next
    
    'WndProc = DefWindowProcA(hWnd, uMsg, wParam, lParam)
End Function

Public Function GetTimerID() As Long
    Dim nID As Long
    Dim T As cTimer
    Dim i As Long
    
    If TimerList.GetCount >= 65535 Then Exit Function
    
    Randomize
    
NewID:
    nID = Int(65535 * Rnd + 1)
    
    If TimerList.GetCount > 0 Then
        For i = 0 To TimerList.GetCount - 1
            Set T = TimerList.GetObject(i)
                If T.ID = nID Then
                    Set T = Nothing
                    GoTo NewID
                End If
            Set T = Nothing
        Next
    End If
    
    GetTimerID = nID
End Function

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal EventID As Long, ByVal dwTime As Long)
    If TimerList.GetCount = 0 Then Exit Sub
    
    Dim T As cTimer
    Dim i As Long
    
    For i = 0 To TimerList.GetCount - 1
        Set T = TimerList.GetObject(i)
            If T.ID = EventID Then
                T.CallOnTime
                Set T = Nothing
                Exit For
            End If
        Set T = Nothing
    Next
End Sub

Public Function LowWord(ByVal inDWord As Long) As Integer
    LowWord = inDWord And &H7FFF&
    If (inDWord And &H8000&) Then LowWord = LowWord Or &H8000
End Function

Public Function HighWord(ByVal inDWord As Long) As Integer
    HighWord = LowWord(((inDWord And &HFFFF0000) \ &H10000) And &HFFFF&)
End Function