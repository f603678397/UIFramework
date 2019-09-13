VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ImageList       As New cObjectList
Private KeyList         As New cArrayList

Private Sub Class_Terminate(): Release: End Sub

Friend Sub Release(): ImageList.Clear: End Sub

Public Function LoadImage(ByVal imgPath As String, ByVal Key As String) As Boolean
    Dim Img As New cImage
    
    If Dir(imgPath) = "" Or Trim(Key) = "" Then Exit Function
    
    Img.FromFile imgPath
    If Img.Handle = 0 Then Exit Function
    
    ImageList.Add Img
    KeyList.Add LCase(Trim(Key))
    LoadImage = True
End Function

Private Sub AddImage(ByVal imgPath As String, ByVal Key As String)
    Dim Img As New cImage
    
    Img.FromFile imgPath
    If Img.Handle = 0 Then Exit Sub
    
    ImageList.Add Img
    KeyList.Add LCase(Trim(Key))
End Sub

Public Sub LoadAllImage(ByVal DirPath As String)
    Dim FileName As String
    Dim Key As String
    
    DirPath = Trim(DirPath)
    If DirPath = "" Then Exit Sub
    DirPath = IIf(Right(DirPath, 1) = "\", DirPath, DirPath & "\")
    If Dir(DirPath, vbDirectory) = "" Then Exit Sub
    
    FileName = Dir(DirPath & "*.*")
    
    Do While FileName <> ""
        Key = Left$(FileName, InStrRev(FileName, ".") - 1)
        AddImage DirPath & FileName, Key
        FileName = Dir()
    Loop
End Sub

Public Function GetImage(ByVal Key As String) As cImage
    Dim Position As Long
    
    Position = KeyList.Find(LCase(Trim(Key)))
    If Position = -1 Then Exit Function
    Set GetImage = ImageList.GetObject(Position)
End Function

Public Sub ReplaceImage(ByVal Key As String, NewImage As cImage)
    Dim Position As Long
    
    Position = KeyList.Find(LCase(Trim(Key)))
    If Position = -1 Then Exit Sub
    
    ImageList.SetObject NewImage, Position
End Sub

Public Sub RemoveImage(ByVal Key As String)
    Dim Position As Long
    Position = KeyList.Find(LCase(Trim(Key)))
    If Position = -1 Then Exit Sub
    ImageList.Remove Position
    KeyList.Remove Position
End Sub

Public Function GetCount() As Long: GetCount = ImageList.GetCount: End Function
Public Function Clear(): ImageList.Clear: KeyList.Clear: End Function