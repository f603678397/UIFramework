VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private ArrayList()     As Variant
Private ArrayCount      As Long

Private Sub Class_Terminate()
    Clear
End Sub

Public Sub Add(T As Variant, Optional Position As Long = -1)
    ArrayCount = ArrayCount + 1
    ReDim Preserve ArrayList(ArrayCount - 1) As Variant
    
    If Position <= -1 Or Position >= ArrayCount - 1 Then
        ArrayList(ArrayCount - 1) = T
    Else
        For i = (ArrayCount - 1) To (Position + 1) Step -1
            ArrayList(i) = ArrayList(i - 1)
        Next
        ArrayList(Position) = T
    End If
End Sub

Public Sub Remove(Position As Long)
    If Position < 0 Or Position > ArrayCount - 1 Then Exit Sub
    
    For i = Position To ArrayCount - 2
        ArrayList(i) = ArrayList(i + 1)
    Next
    
    ArrayCount = ArrayCount - 1
    If ArrayCount > 0 Then
        ReDim Preserve ArrayList(ArrayCount - 1) As Variant
    Else
        Erase ArrayList
    End If
End Sub

Public Sub Clear()
    Erase ArrayList
    ArrayCount = 0
End Sub

Public Sub SetValue(T As Variant, Position As Long)
    If Position < 0 Or Position > ArrayCount - 1 Then Exit Sub
    ArrayList(Position) = T
End Sub

Public Function GetValue(Position As Long) As Variant
    If Position < 0 Or Position > ArrayCount - 1 Then Exit Function
    GetValue = ArrayList(Position)
End Function

Public Function Find(T As Variant) As Long
    Dim Position  As Long
    Position = -1
    
    For i = 0 To ArrayCount - 1
        If ArrayList(i) = T Then
            Position = i
            Exit For
        End If
    Next
    
    Find = Position
End Function

Public Function GetCount() As Long
    GetCount = ArrayCount
End Function
