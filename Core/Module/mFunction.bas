Attribute VB_Name = "mFunction"
Option Explicit
Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WndProc = DefWindowProcA(hwnd, uMsg, wParam, lParam)
End Function

Public Function GetAddress(ByVal lngAddr As Long) As Long
    GetAddress = lngAddr
End Function
