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

Private Sub Class_Terminate(): DisposeAll: End Sub

Public Function LoadImage(ByVal imgPath As String, ByVal Key As String) As Boolean
    If Dir(imgPath) = "" Or Trim(Key) = "" Then Exit Function
    LoadImage = AddImage(imgPath, Key)
End Function

Private Function AddImage(ByVal imgPath As String, ByVal Key As String) As Boolean
    Dim Img As New cImage
    
    Img.FromFile imgPath
    If Img.Handle = 0 Then Exit Function
    
    ImageList.Add Img
    KeyList.Add LCase(Trim(Key))
    AddImage = True
End Function

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

Public Sub DisposeImage(ByVal Key As String)
    Dim Position As Long
    Dim Img As cImage
    
    Position = KeyList.Find(LCase(Trim(Key)))
    If Position = -1 Then Exit Sub
    
    Set Img = ImageList.GetObject(Position)
    Img.Release
    ImageList.Remove Position
    KeyList.Remove Position
End Sub

Public Sub DisposeAll()
    Dim i As Long
    Dim Img As cImage
    
    For i = 0 To ImageList.GetCount - 1
        Set Img = ImageList.GetObject(i)
            Img.Release
        Set Img = Nothing
    Next
    ImageList.Clear
    KeyList.Clear
End Sub

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
