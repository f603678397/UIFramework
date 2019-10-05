Attribute VB_Name = "mFunction"
Option Explicit

Public ImageManager     As cImageManager

Public Type CursorType
    Normal      As Long
    All         As Long
    NS          As Long
    WE          As Long
    NE_SW       As Long
    NW_SE       As Long
End Type

Public Cursor   As CursorType

Public Sub Main()
    With Cursor
        .Normal = LoadCursorWL(ByVal 0&, IDC_ARROW)
        .All = LoadCursorWL(ByVal 0&, IDC_SIZEALL)
        .NS = LoadCursorWL(ByVal 0&, IDC_SIZENS)
        .WE = LoadCursorWL(ByVal 0&, IDC_SIZEWE)
        .NE_SW = LoadCursorWL(ByVal 0&, IDC_SIZENESW)
        .NW_SE = LoadCursorWL(ByVal 0&, IDC_SIZENWSE)
    End With

    cCore.Initialize
    Set ImageManager = cCore.GetImageManager
    frmSplash.Show
End Sub

Public Sub ExitApp()
    cCore.Terminate
End Sub

