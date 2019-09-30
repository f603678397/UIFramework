VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrawable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Bitmap      As cBitmap
Private Graphics    As cGraphics
Private Folder      As String

Private Type LineParam
    X1      As Single
    Y1      As Single
    X2      As Single
    Y2      As Single
    Edge    As Single
    Color   As Long
End Type

Private Type ArcParam
    X       As Single
    Y       As Single
    Width   As Single
    Height  As Single
    Start   As Single
    Angle   As Single
    Edge    As Single
    Color   As Long
End Type

Private Type RectParam
    X       As Single
    Y       As Single
    Width   As Single
    Height  As Single
    Edge    As Single
    Color   As Long
End Type

Private Type OvalParam
    X       As Single
    Y       As Single
    Width   As Single
    Height  As Single
    Edge    As Single
    Color   As Long
End Type

Private Type StringParam
    Text        As String
    FontName    As String
    FontSize    As Single
    Align       As Drawing2D.StringAlignment
    Bound       As Drawing2D.RECTF
    Color       As Long
End Type

Private Sub Class_Initialize(): Release: End Sub

Friend Sub Release()
    Folder = ""
    Set Graphics = Nothing
    Set Bitmap = Nothing
    DrawableList.RemoveObject Me
End Sub

Public Function LoadFromXML(ByVal xmlPath As String) As Boolean
    Dim xmlDoc  As New DOMDocument60
    Dim Node    As IXMLDOMNode
    
    If Dir(xmlPath) = vbNullString Then Exit Function
    
    If Not xmlDoc.Load(xmlPath) Then Exit Function
    Folder = Left$(xmlPath, InStrRev(xmlPath, "\"))
    
    Set Node = xmlDoc.selectSingleNode("canvas")
    If Node Is Nothing Then Exit Function
    
    CreateCanvas Node
    DrawShape Node

    DrawableList.Add Me
    LoadFromXML = True
End Function

Public Function ToImage() As cImage
    Set ToImage = Bitmap.ToImage
End Function

Private Sub CreateCanvas(Node As IXMLDOMNode)
    Dim NodeMap     As IXMLDOMNamedNodeMap
    Dim NodeName    As String
    Dim NodeValue   As String
    
    Dim Width       As Integer, Height As Integer
    Dim Color       As Long
    Dim i           As Integer
'----------------------------------------------------------------------------
    Color = cColor.Transparent
    
    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Sub
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        
        Select Case NodeName
            Case LCase("Width")
                Width = NodeValue
            Case LCase("Height")
                Height = NodeValue
            Case LCase("Color")
                Color = ColorToLong(NodeValue)
        End Select
    Next
    
    Set Graphics = Nothing
    Set Bitmap = Nothing
    Set Bitmap = New cBitmap

    Bitmap.CreateEmptyBitmap Width, Height
    Set Graphics = Bitmap.GetGraphics
    Graphics.SetSmoothingMode SmoothingModeAntiAlias
    
    Graphics.Clear Color
End Sub

Private Sub DrawShape(ParentNode As IXMLDOMNode)
    Dim Node    As IXMLDOMNode
    Dim i       As Integer
'----------------------------------------------------------------------------
    For i = 0 To ParentNode.childNodes.Length - 1
        Set Node = ParentNode.childNodes.Item(i)
            Select Case LCase(Node.NodeName)
                Case LCase("Line")
                    DrawLine Node
                Case LCase("Arc")
                    DrawArc Node
                Case LCase("Rect")
                    DrawRect Node
                Case LCase("Oval")
                    DrawOval Node
                Case LCase("Img")
                    DrawImg Node
                Case LCase("String")
                    DrawString Node
                Case LCase("Path")
                    DrawPath Node
            End Select
        Set Node = Nothing
    Next
End Sub

Private Function ColorToLong(ByVal Value As String) As Long
    On Error GoTo Default
    ColorToLong = CLng(Replace(Value, "#", "&H"))
    Exit Function
Default:
    ColorToLong = cColor.Transparent
End Function

Private Function GetStringParam(Node As IXMLDOMNode, ByRef Out As StringParam) As Boolean
    Dim NodeMap     As IXMLDOMNamedNodeMap
    Dim NodeName    As String
    Dim NodeValue   As String
    Dim D()         As String
    Dim i           As Integer
    
    Dim Color       As Long
'----------------------------------------------------------------------------
    Color = cColor.Black
    
    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Function
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        Select Case NodeName
            Case LCase("Text")
                Out.Text = Replace(NodeValue, "\n", vbCrLf)
            Case LCase("Position")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Out.Bound.Left = CSng(D(0))
                Out.Bound.Top = CSng(D(1))
            Case LCase("Size")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Out.Bound.Right = CSng(D(0))
                Out.Bound.bottom = CSng(D(1))
            Case LCase("Align")
                Select Case LCase(NodeValue)
                    Case LCase("Left")
                        Out.Align = StringAlignmentNear
                    Case LCase("Center")
                        Out.Align = StringAlignmentCenter
                    Case LCase("Right")
                        Out.Align = StringAlignmentFar
                    Case Else
                        Out.Align = StringAlignmentNear
                End Select
            Case LCase("FontName")
                Out.FontName = NodeValue
            Case LCase("FontSize")
                Out.FontSize = CSng(NodeValue)
            Case LCase("Color")
                Color = ColorToLong(NodeValue)
        End Select
    Next
    
    Out.Color = Color
    GetStringParam = True
End Function

Private Function GetLineParam(Node As IXMLDOMNode, ByRef Out As LineParam) As Boolean
    Dim NodeMap     As IXMLDOMNamedNodeMap
    Dim NodeName    As String
    Dim NodeValue   As String
    Dim D()         As String
    Dim i           As Integer
    
    Dim X1          As Single, Y1       As Single
    Dim X2          As Single, Y2       As Single
    Dim Edge        As Single, Color    As Long
'----------------------------------------------------------------------------
    Color = cColor.Black
    Edge = 1
    
    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Function
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        Select Case NodeName
            Case LCase("Start")
                D = Split(Replace(NodeValue, " ", ""), ",")
                X1 = CSng(D(0))
                Y1 = CSng(D(1))
            Case LCase("End")
                D = Split(Replace(NodeValue, " ", ""), ",")
                X2 = CSng(D(0))
                Y2 = CSng(D(1))
            Case LCase("Edge")
                Edge = CSng(NodeValue)
            Case LCase("Color")
                Color = ColorToLong(NodeValue)
        End Select
    Next
    
    With Out
        .X1 = X1: .Y1 = Y1
        .X2 = X2: .Y2 = Y2
        .Edge = Edge: .Color = Color
    End With
    
    GetLineParam = True
End Function

Private Function GetArcParam(Node As IXMLDOMNode, ByRef Out As ArcParam) As Boolean
    Dim NodeMap     As IXMLDOMNamedNodeMap
    Dim NodeName    As String
    Dim NodeValue   As String
    Dim D()         As String
    Dim i           As Integer
    
    Dim X           As Single, Y        As Single
    Dim Width       As Single, Height   As Single
    Dim Start       As Single, Angle    As Single
    Dim Edge        As Single, Color    As Long
'----------------------------------------------------------------------------
    Color = cColor.Black
    Edge = 1
    
    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Function
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        Select Case NodeName
            Case LCase("Position")
                D = Split(Replace(NodeValue, " ", ""), ",")
                X = CSng(D(0))
                Y = CSng(D(1))
            Case LCase("Size")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Width = CSng(D(0))
                Height = CSng(D(1))
            Case LCase("Start")
                Start = CSng(NodeValue)
            Case LCase("Angle")
                Angle = CSng(NodeValue)
            Case LCase("Edge")
                Edge = CSng(NodeValue)
            Case LCase("Color")
                Color = ColorToLong(NodeValue)
        End Select
    Next
    
    With Out
        .X = X: .Y = Y
        .Width = Width: .Height = Height
        .Start = Start: .Angle = Angle
        .Edge = Edge: .Color = Color
    End With
    
    GetArcParam = True
End Function

Private Function GetRectParam(Node As IXMLDOMNode, ByRef Out As RectParam) As Boolean
    Dim NodeMap     As IXMLDOMNamedNodeMap
    Dim NodeName    As String
    Dim NodeValue   As String
    Dim D()         As String
    Dim i           As Integer
    
    Dim X           As Single, Y        As Single
    Dim Width       As Single, Height   As Single
    Dim Edge        As Single, Color    As Long
'----------------------------------------------------------------------------
    Color = cColor.Black
    Edge = 1
    
    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Function
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        Select Case NodeName
            Case LCase("Position")
                D = Split(Replace(NodeValue, " ", ""), ",")
                X = CSng(D(0))
                Y = CSng(D(1))
            Case LCase("Size")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Width = CSng(D(0))
                Height = CSng(D(1))
            Case LCase("Edge")
                Edge = CSng(NodeValue)
            Case LCase("Color")
                Color = ColorToLong(NodeValue)
        End Select
    Next
    
    With Out
        .X = X: .Y = Y
        .Width = Width: .Height = Height
        .Edge = Edge: .Color = Color
    End With
    
    GetRectParam = True
End Function

Private Function GetOvalParam(Node As IXMLDOMNode, ByRef Out As OvalParam) As Boolean
    Dim Param As RectParam
    If Not GetRectParam(Node, Param) Then Exit Function
    
    With Out
        .X = Param.X: .Y = Param.Y
        .Width = Param.Width: .Height = Param.Height
        .Edge = Param.Edge: .Color = Param.Color
    End With
    
    GetOvalParam = True
End Function

Private Function GetBrush(ParentNode As IXMLDOMNode) As cBrush
    Dim Node    As IXMLDOMNode
    Dim i       As Integer
    
    For i = 0 To ParentNode.childNodes.Length - 1
        Set Node = ParentNode.childNodes.Item(i)
            Select Case LCase(Node.NodeName)
                Case LCase("Solid")
                    Set GetBrush = GetSolidBrush(Node)
                    Set Node = Nothing
                    Exit For
                Case LCase("Gradient")
                    Set GetBrush = GetGradientBrush(Node)
                    Set Node = Nothing
                    Exit For
            End Select
        Set Node = Nothing
    Next
End Function

Private Function GetSolidBrush(Node As IXMLDOMNode) As cBrush
    Dim Brush       As New cSolidBrush
    Dim NodeMap     As IXMLDOMNamedNodeMap
    Dim NodeName    As String
    Dim NodeValue   As String
    Dim Color       As Long
    Dim i           As Integer
    
    Color = cColor.Black
    
    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Function
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        
        If NodeName = LCase("Color") Then
            Color = ColorToLong(NodeValue)
        End If
    Next
    
    Brush.Color = Color
    
    Set GetSolidBrush = Brush.GetBaseBrush
End Function

Private Function GetGradientBrush(Node As IXMLDOMNode) As cBrush
    Dim Brush       As New cLinearBrush
    Dim NodeMap     As IXMLDOMNamedNodeMap
    Dim NodeName    As String
    Dim NodeValue   As String
    Dim Point1      As POINTF
    Dim Point2      As POINTF
    Dim Color1      As Long
    Dim Color2      As Long
    Dim D()         As String
    Dim i           As Integer
    
    Color1 = cColor.Transparent
    Color2 = cColor.Transparent
    
    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Function
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        
        Select Case NodeName
            Case LCase("Position")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Point1.X = D(0): Point1.Y = D(1)
            Case LCase("Vector")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Point2.X = D(0): Point2.Y = D(1)
            Case LCase("Start")
                Color1 = ColorToLong(NodeValue)
            Case LCase("End")
                Color2 = ColorToLong(NodeValue)
        End Select
    Next
    
    Brush.SetStartPoint Point1.X, Point1.Y
    Brush.SetEndPoint Point2.X, Point2.Y
    Brush.SetColors Color1, Color2
    Brush.FillWrapMode = WrapModeTileFlipXY
    
    Set GetGradientBrush = Brush.GetBaseBrush
End Function

Private Sub DrawLine(Node As IXMLDOMNode)
    Dim Param   As LineParam
    Dim Pen     As New cPen
'----------------------------------------------------------------------------
    If Not GetLineParam(Node, Param) Then Exit Sub
    With Param
        Pen.Color = .Color
        Pen.Width = .Edge
        Graphics.DrawLine Pen, .X1, .Y1, .X2, .Y2
    End With
End Sub

Private Sub DrawArc(Node As IXMLDOMNode)
    Dim Param   As ArcParam
    Dim Pen     As New cPen
'----------------------------------------------------------------------------
    If Not GetArcParam(Node, Param) Then Exit Sub
    
    With Param
        Pen.Color = .Color
        Pen.Width = .Edge
        Graphics.DrawArc Pen, .X, .Y, .Width, .Height, .Start, .Angle
    End With
End Sub

Private Sub DrawRect(Node As IXMLDOMNode)
    Dim Param   As RectParam
    Dim Pen     As New cPen
    Dim Brush   As cBrush
    
    If Not GetRectParam(Node, Param) Then Exit Sub
    Set Brush = GetBrush(Node)
    
    With Param
        If Not Brush Is Nothing Then
            Graphics.FillRectangle Brush, .X, .Y, .Width, .Height
        End If
        Pen.Color = .Color
        Pen.Width = .Edge
        Graphics.DrawRectangle Pen, .X, .Y, .Width, .Height
    End With
End Sub

Private Sub DrawOval(Node As IXMLDOMNode)
    Dim Param   As OvalParam
    Dim Pen     As New cPen
    Dim Brush   As cBrush
    
    If Not GetOvalParam(Node, Param) Then Exit Sub
    Set Brush = GetBrush(Node)
    
    With Param
        If Not Brush Is Nothing Then
            Graphics.FillEllipse Brush, .X, .Y, .Width, .Height
        End If
        Pen.Color = .Color
        Pen.Width = .Edge
        Graphics.DrawEllipse Pen, .X, .Y, .Width, .Height
    End With
End Sub

Private Sub DrawImg(Node As IXMLDOMNode)
    Dim NodeMap         As IXMLDOMNamedNodeMap
    Dim NodeName        As String
    Dim NodeValue       As String
    
    Dim i       As Integer
    Dim D()     As String
    Dim Path    As String
    Dim Dst     As Drawing2D.RECTF
    Dim Src     As Drawing2D.RECTF
    Dim Begin   As String
    Dim Img     As New cImage
    
    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Sub
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        Select Case NodeName
            Case LCase("Src")
                NodeValue = Replace(NodeValue, "/", "\")
                NodeValue = Replace(NodeValue, "\/", "\")
                Begin = Left$(NodeValue, InStr(NodeValue, "\"))
                If Begin = ".\" Then
                    Path = Folder & Right$(NodeValue, Len(NodeValue) - InStr(NodeValue, "\"))
                Else
                    Path = NodeValue
                End If
            Case LCase("Position")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Dst.Left = D(0)
                Dst.Top = D(1)
            Case LCase("Size")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Dst.Right = D(0)
                Dst.bottom = D(1)
            Case LCase("Clip")
                D = Split(Replace(NodeValue, " ", ""), ",")
                Src.Left = D(0): Src.Top = D(1)
                Src.Right = D(2): Src.bottom = D(3)
        End Select
    Next

    Img.FromFile Path
    
    If Src.Left = 0 And Src.Top = 0 And Src.Right = 0 And Src.bottom = 0 Then
        Src.Right = Img.Width
        Src.bottom = Img.Handle
    End If
    Graphics.DrawImageRectRect Img, Dst, Src, Nothing
End Sub

Private Sub DrawString(Node As IXMLDOMNode)
    Dim Param       As StringParam
    Dim Font        As New cFont
    Dim Format      As New cStringFormat
    Dim Family      As New cFontFamily
    Dim Brush       As New cSolidBrush
'----------------------------------------------------------------------------
    If Not GetStringParam(Node, Param) Then Exit Sub
    
    With Param
        Brush.Color = .Color
        Family.FromName .FontName
        Font.FromFamily Family, .FontSize, FontStyleRegular, UnitPixel
        Format.Align = .Align

        Graphics.DrawString .Text, Font, .Bound, Format, Brush.GetBaseBrush
    End With
End Sub

Private Sub DrawPath(Node As IXMLDOMNode)
    Dim NodeMap     As IXMLDOMNamedNodeMap
    Dim NodeName    As String
    Dim NodeValue   As String
    
    Dim i           As Long
    Dim Path        As New cPath
    Dim Width       As Single
    Dim Color       As Long
    Dim bClose      As Boolean
    Dim Pen         As New cPen
    Dim Brush       As cBrush
'----------------------------------------------------------------------------
    Color = cColor.Black
    Width = 1
    bClose = True

    Set NodeMap = Node.Attributes
    If NodeMap Is Nothing Then Exit Sub
    
    For i = 0 To NodeMap.Length - 1
        NodeName = LCase(NodeMap.Item(i).NodeName)
        NodeValue = NodeMap.Item(i).NodeValue
        Select Case NodeName
            Case LCase("Edge")
                Width = CSng(NodeValue)
            Case LCase("Color")
                Color = ColorToLong(NodeValue)
            Case LCase("Close")
                bClose = CBool(NodeValue)
        End Select
    Next
    
    AddPath Path, Node
    Set Brush = GetBrush(Node)
    If bClose Then Path.CloseFigure
    
    If Not Brush Is Nothing Then
        Graphics.FillPath Brush, Path
    End If
    Graphics.DrawPath Path, Pen
End Sub

Private Sub AddPath(Path As cPath, ParentNode As IXMLDOMNode)
    Dim Node    As IXMLDOMNode
    Dim i       As Integer
'----------------------------------------------------------------------------
    For i = 0 To ParentNode.childNodes.Length - 1
        Set Node = ParentNode.childNodes.Item(i)
            Select Case LCase(Node.NodeName)
                Case LCase("Line")
                    AddLine Path, Node
                Case LCase("Arc")
                    AddArc Path, Node
                Case LCase("Rect")
                    AddRect Path, Node
                Case LCase("Oval")
                    AddOval Path, Node
                Case LCase("String")
                    AddString Path, Node
            End Select
        Set Node = Nothing
    Next
End Sub

Private Sub AddString(Path As cPath, Node As IXMLDOMNode)
    Dim Param       As StringParam
    Dim Format      As New cStringFormat
    Dim Family      As New cFontFamily
'----------------------------------------------------------------------------
    If Not GetStringParam(Node, Param) Then Exit Sub
    
    With Param
        Family.FromName .FontName
        Format.Align = .Align
        
        Path.AddString .Text, Family, FontStyleRegular, .FontSize, .Bound, Format
    End With
End Sub

Private Sub AddLine(Path As cPath, Node As IXMLDOMNode)
    Dim Param   As LineParam
'----------------------------------------------------------------------------
    If Not GetLineParam(Node, Param) Then Exit Sub
    
    With Param
        Path.AddLine .X1, .Y1, .X2, .Y2
    End With
End Sub

Private Sub AddArc(Path As cPath, Node As IXMLDOMNode)
    Dim Param   As ArcParam
'----------------------------------------------------------------------------
    If Not GetArcParam(Node, Param) Then Exit Sub
    
    With Param
        Path.AddArc .X, .Y, .Width, .Height, .Start, .Angle
    End With
End Sub

Private Sub AddRect(Path As cPath, Node As IXMLDOMNode)
    Dim Param   As RectParam
    If Not GetRectParam(Node, Param) Then Exit Sub
    
    With Param
        Path.AddRectangle .X, .Y, .Width, .Height
    End With
End Sub

Private Sub AddOval(Path As cPath, Node As IXMLDOMNode)
    Dim Param   As OvalParam
    If Not GetOvalParam(Node, Param) Then Exit Sub
    
    With Param
        Path.AddEllipse .X, .Y, .Width, .Height
    End With
End Sub
