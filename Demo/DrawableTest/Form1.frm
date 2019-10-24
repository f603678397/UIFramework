VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "DrawableTest"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Activity                As cActivity
Dim Layout                  As cLayout
Dim WithEvents View         As cView
Attribute View.VB_VarHelpID = -1
Dim Drawable                As New cDrawable

Private Sub Form_Load()
    cCore.Initialize
    Set Activity = cCore.CreateActivity(Me.hWnd)
    Set Layout = Activity.CreateLayout
    Activity.SetLayout Layout
    
    Set View = Layout.CreateView(0, 0, 800, 600)
    
    Drawable.LoadFromXML App.Path & "\Res\xml\Ä£°å.xml"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cCore.Terminate
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Canvas.DrawImage Drawable.ToImage, 0, 0
End Sub
