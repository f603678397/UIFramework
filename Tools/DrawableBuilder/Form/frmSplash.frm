VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Splash"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "frmSplash"
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
Public WithEvents Timer1    As cTimer
Attribute Timer1.VB_VarHelpID = -1

Private Sub Form_Load()
    Drawable.LoadFromXML App.Path & "\res\splash.xml"
    
    Set Activity = cCore.CreateActivity(Me.hWnd, True)
    Set Layout = Activity.CreateLayout
    Set View = Layout.CreateView(0, 0, 301, 151)
    Activity.SetLayout Layout
    
    Set Timer1 = New cTimer
    Timer1.Create Me.hWnd
    Timer1.Interval = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Timer1 = Nothing
    Set View = Nothing
    Set Layout = Nothing
    Set Drawable = Nothing
    cCore.DestroyActivity Activity
End Sub

Private Sub Timer1_onTime()
    frmMain.Show
    Unload Me
End Sub

Private Sub View_Paint(Canvas As Drawing2D.cGraphics)
    Canvas.DrawImage Drawable.GetImage, 0, 0
End Sub
