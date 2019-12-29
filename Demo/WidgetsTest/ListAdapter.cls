VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cListAdapter

Dim Item(19) As String

Private Sub Class_Initialize()
    Dim i As Long
    For i = 0 To UBound(Item)
        Item(i) = "Item_" & Format(i + 1, "00")
    Next
End Sub

Private Function cListAdapter_GetCount() As Long
    cListAdapter_GetCount = UBound(Item) + 1
End Function

Private Sub cListAdapter_DrawItem(Canvas As Drawing2D.cGraphics, ByVal nIndex As Long, ByVal State As ItemState)
    Dim Font        As New cFont
    Dim Fam         As New cFontFamily
    Dim strFormat   As New cStringFormat

    Dim Brush       As New cSolidBrush
    Dim Bound       As RECTF
    Dim bgColor     As Long

    Fam.FromName "Î¢ÈíÑÅºÚ"
    Font.FromFamily Fam, 12, FontStyleRegular, UnitPixel
    strFormat.Align = StringAlignmentNear
    strFormat.LineAlign = StringAlignmentCenter
    strFormat.Flags = StringFormatFlagsNoClip

    Bound.Right = 120
    Bound.Bottom = Form1.ListView.GetItemHeight

    bgColor = Form1.ListView.BackColor
    
    Brush.Color = cColor.FromARGB(255, _
                                  255 - cColor.GetRed(bgColor), _
                                  255 - cColor.GetGreen(bgColor), _
                                  255 - cColor.GetBlue(bgColor))
    
    Select Case State
        Case ItemState.IS_NORMAL
            Canvas.Clear bgColor
        Case ItemState.IS_HOVER
            Canvas.Clear cColor.FromARGB(255, 30, 110, 195)
        Case ItemState.IS_DOWN
            Canvas.Clear cColor.FromARGB(255, 24, 88, 156)
    End Select
    Canvas.DrawString Item(nIndex), Font, Bound, strFormat, Brush.GetBaseBrush
End Sub


