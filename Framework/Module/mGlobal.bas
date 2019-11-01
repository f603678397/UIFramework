Attribute VB_Name = "mGlobal"
Option Explicit

Public Drawing              As New cDrawing
Public TimerList            As New cObjectList
Public ActivityList         As New cObjectList
Public IM                   As New cImageManager
Public DrawableList         As New cObjectList

' Debug ¿ª¹Ø
Public bShowRedrawRgn       As Boolean

Public Type RECTI
    Left    As Integer
    Top     As Integer
    Width   As Integer
    Height  As Integer
End Type

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
    
    WndProc = DefWindowProcA(hWnd, uMsg, wParam, lParam)
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

Public Function GetButton(ByVal n As Long) As Integer
    Dim Button As Integer
    
    If (n And MK_LBUTTON) = MK_LBUTTON Then Button = Button Or vbLeftButton
    If (n And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or vbRightButton
    If (n And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or vbMiddleButton
    
    GetButton = Button
End Function

Public Function GetShift(ByVal n As Long) As Integer
    Dim Shift As Integer
    
    If (n And MK_SHIFT) = MK_SHIFT Then Shift = Shift Or &H1
    If (n And MK_CONTROL) = MK_CONTROL Then Shift = Shift Or &H2
    
    GetShift = Shift
End Function
