VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Graphics As cGraphics
Dim sText   As String

Private Sub Command2_Click()
    Dim Fam As New cFontFamily
    Dim Font As New cFont
    Dim StrFormat As New cStringFormat
    Dim Ranges() As CharacterRange
    Dim Rgn() As cRegion
    
    Dim Brush As New cSolidBrush
    Dim Pen As New cPen
    Dim ts(0 To 3) As Single
    
    Dim Layout As RECTF
    Dim Count As Long
    Dim i As Long
    Dim RF() As RECTF
    
    Count = Len(sText)
    
    ReDim Ranges(Count - 1) As CharacterRange
    ReDim RF(Count - 1) As RECTF
    
    For i = 0 To Count - 1
        Ranges(i).First = i
        Ranges(i).Length = 1
    Next
    
    For i = 0 To UBound(ts)
        ts(i) = 30
    Next
    
    With Layout
        .Top = 0
        .Left = 0
        .Right = Me.ScaleWidth
        .Bottom = Me.ScaleHeight
    End With
    
    Fam.FromName "Î¢ÈíÑÅºÚ"
    Font.FromFamily Fam, 12, FontStyleRegular, UnitPixel
'    StrFormat.GenericTypographic
    StrFormat.Flags = StringFormatFlagsNoClip Or StringFormatFlagsMeasureTrailingSpaces
    StrFormat.SetMeasurableCharacterRanges Ranges
    StrFormat.SetTabStops 0, ts
    
    Pen.Color = cColor.Red
    
    Rgn = Graphics.MeasureCharacterRanges(sText, Font, Layout, StrFormat)
    
    For i = 0 To Count - 1
        RF(i) = Rgn(i).GetBounds(Graphics)
    Next
    
    'Me.Caption = UBound(RF) + 1 & " " & RF(10).Left & " " & RF(10).Top
    
    Graphics.Clear cColor.White
    Graphics.SetTextRenderingHint TextRenderingHintClearTypeGridFit
    Graphics.DrawString sText, Font, Layout, StrFormat, Brush.GetBaseBrush
    Graphics.DrawRectangles Pen, RF
End Sub

Private Sub Form_Load()
    cDrawing.Init
    Set Graphics = cDrawing.CreateGraphicsFromHDC(Me.hDC)
    
    sText = vbLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Graphics = Nothing
    cDrawing.Shutdown
End Sub
