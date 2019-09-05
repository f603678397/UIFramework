VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private ObjectList()        As Object
Private ObjectCount         As Long

Private Sub Class_Terminate()
    Clear
End Sub

Public Sub Add(T As Object, Optional Position As Long = -1)
    If T Is Nothing Then Exit Sub
    ObjectCount = ObjectCount + 1
    ReDim Preserve ObjectList(ObjectCount - 1) As Object
    
    If Position <= -1 Or Position >= ObjectCount - 1 Then
        Set ObjectList(ObjectCount - 1) = T
    Else
        For i = (ObjectCount - 1) To (Position + 1) Step -1
            Set ObjectList(i) = ObjectList(i - 1)
            Set ObjectList(i - 1) = Nothing
        Next
        Set ObjectList(Position) = T
    End If
End Sub

Public Sub AddList(List As cObjectList)
    If List Is Nothing Then Exit Sub
    
    Dim i As Long
    Dim Obj As Object
    
    For i = 0 To List.GetCount - 1
        Set Obj = List.GetObject(i)
        Add Obj
        Set Obj = Nothing
    Next
End Sub

Public Sub MoveObject(T As Object, ByVal NewPosition As Long)
    Dim Position As Long
    Dim S As Long, E As Long
    
    Position = Find(T)
    If Position = -1 Then Exit Sub
    If Position = NewPosition Then Exit Sub
    If NewPosition < 0 Or NewPosition > ObjectCount - 1 Then Exit Sub
    
    If NewPosition < Position Then
        S = NewPosition
        E = Position
    Else
        S = Position
        E = NewPosition
    End If
    
    For i = E To S + 1 Step -1
        Set ObjectList(i) = ObjectList(i - 1)
    Next
    Set ObjectList(S) = T
End Sub

Public Sub RemoveObject(T As Object)
    Remove Find(T)
End Sub

Public Sub Remove(Position As Long)
    If Position < 0 Or Position > ObjectCount - 1 Then Exit Sub
    
    For i = Position To ObjectCount - 2
        Set ObjectList(i) = Nothing
        Set ObjectList(i) = ObjectList(i + 1)
    Next
    
    Set ObjectList(ObjectCount - 1) = Nothing
    ObjectCount = ObjectCount - 1
    If ObjectCount > 0 Then
        ReDim Preserve ObjectList(ObjectCount - 1) As Object
    Else
        Erase ObjectList
    End If
End Sub

Public Sub Clear()
    For i = 0 To ObjectCount - 1
        Set ObjectList(i) = Nothing
    Next
    
    Erase ObjectList
    ObjectCount = 0
End Sub

Public Sub SetObject(T As Object, Position As Long)
    If Position < 0 Or Position > ObjectCount - 1 Then Exit Sub
    Set ObjectList(Position) = Nothing
    Set ObjectList(Position) = T
End Sub

Public Function GetObject(Position As Long) As Object
    If Position < 0 Or Position > ObjectCount - 1 Then Exit Function
    Set GetObject = ObjectList(Position)
End Function

Public Function Find(T As Object) As Long
    Dim Position  As Long
    Position = -1

    For i = 0 To ObjectCount - 1
        If ObjectList(i) Is T Then
            Position = i
            Exit For
        End If
    Next

    Find = Position
End Function

Public Function GetCount() As Long
    GetCount = ObjectCount
End Function
