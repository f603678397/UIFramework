VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long): mHandle = nHandle: SetPageUnit GlobalUnit: End Property

Public Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeleteGraphics mHandle
    mHandle = 0
End Sub

Public Function GetDC() As Long
    GdipGetDC mHandle, GetDC
End Function

Public Sub ReleaseDC(ByVal hdc As Long)
    GdipReleaseDC mHandle, hdc
End Sub

Public Sub SetCompositingMode(Mode As CompositingMode)
    GdipSetCompositingMode mHandle, Mode
End Sub

Public Function GetCompositingMode() As CompositingMode
    GdipGetCompositingMode mHandle, GetCompositingMode
End Function

Public Sub SetCompositingQuality(Quality As CompositingQuality)
    GdipSetCompositingQuality mHandle, Quality
End Sub

Public Function GetCompositingQuality() As CompositingQuality
    GdipGetCompositingQuality mHandle, GetCompositingQuality
End Function

Public Sub SetSmoothingMode(Mode As SmoothingMode)
    GdipSetSmoothingMode mHandle, Mode
End Sub

Public Function GetSmoothingMode() As SmoothingMode
    GdipGetSmoothingMode mHandle, GetSmoothingMode
End Function

Public Sub SetRenderingOrigin(ByVal X As Long, ByVal Y As Long)
    GdipSetRenderingOrigin mHandle, X, Y
End Sub

Public Function GetRenderingOrigin() As POINTL
    Dim P As POINTL
    Dim X As Long, Y As Long
    GdipGetRenderingOrigin mHandle, X, Y
    P.X = X: P.Y = Y
    GetRenderingOrigin = P
End Function

Public Sub SetTextRenderingHint(Mode As TextRenderingHint)
    GdipSetTextRenderingHint mHandle, Mode
End Sub

Public Function GetTextRenderingHint() As TextRenderingHint
    GdipGetTextRenderingHint mHandle, GetTextRenderingHint
End Function

Public Sub SetTextContrast(ByVal Contrast As Long)
    GdipSetTextContrast mHandle, Contrast
End Sub

Public Function GetTextContrast() As Long
    GdipGetTextContrast mHandle, GetTextContrast
End Function

Public Function GetInterpolationMode() As InterpolationMode
    GdipGetInterpolationMode mHandle, GetInterpolationMode
End Function

Public Sub SetInterpolationMode(emMode As InterpolationMode)
    GdipSetInterpolationMode mHandle, emMode
End Sub

Public Sub SetPixelOffsetMode(Mode As PixelOffsetMode)
    GdipSetPixelOffsetMode mHandle, Mode
End Sub

Public Function GetPixelOffsetMode() As PixelOffsetMode
    GdipGetPixelOffsetMode mHandle, GetPixelOffsetMode
End Function

Public Sub SetClipGraphics(Graphics As cGraphics, Mode As CombineMode)
    GdipSetClipGraphics mHandle, Graphics.Handle, Mode
End Sub

Public Sub SetClip(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal Mode As CombineMode)
    GdipSetClipRect mHandle, X, Y, Width, Height, Mode
End Sub

Public Sub SetClipPath(Path As cPath, ByVal Mode As CombineMode)
    GdipSetClipPath mHandle, Path.Handle, Mode
End Sub

Public Sub SetClipRegion(region As cRegion, ByVal Mode As CombineMode)
    GdipSetClipRegion mHandle, region.Handle, Mode
End Sub

Public Sub ResetClip()
    GdipResetClip mHandle
End Sub

Public Sub SetPageScale(ByVal fScale As Single)
    GdipSetPageScale mHandle, fScale
End Sub

Public Function GetPageScale() As Single
    GdipGetPageScale mHandle, GetPageScale
End Function

Public Sub SetPageUnit(Unit As GpUnit)
    GdipSetPageUnit mHandle, Unit
End Sub

Public Function GetPageUnit() As GpUnit
    GdipGetPageUnit mHandle, GetPageUnit
End Function

Public Function GetDpiX() As Single
    GdipGetDpiX mHandle, GetDpiX
End Function

Public Function GetDpiY() As Single
    GdipGetDpiY mHandle, GetDpiY
End Function

Public Sub SetScale(ByVal sx As Single, ByVal sy As Single, ByVal Order As MatrixOrder)
    GdipScaleWorldTransform mHandle, sx, sy, Order
End Sub

Public Sub SetTranslate(ByVal dx As Single, ByVal dy As Single, ByVal Order As MatrixOrder)
    GdipTranslateWorldTransform mHandle, dx, dy, Order
End Sub

Public Sub SetRotate(ByVal Angle As Single, Order As MatrixOrder)
    GdipRotateWorldTransform mHandle, Angle, Order
End Sub

Public Sub SetWorldTransform(Matrix As cMatrix)
    GdipSetWorldTransform mHandle, Matrix.Handle
End Sub

Public Function GetWorldTransform() As cMatrix
    Dim Matrix As New cMatrix
    Dim H As Long
    Matrix.Release
    GdipGetWorldTransform mHandle, H
    Matrix.Handle = H
    Set GetWorldTransform = Matrix
End Function

Public Sub ResetWorldTransform()
    GdipResetWorldTransform mHandle
End Sub

Public Function Save() As Long
    GdipSaveGraphics mHandle, Save
End Function

Public Sub Restore(ByVal state As Long)
    GdipRestoreGraphics mHandle, state
End Sub

Public Function BeginContainer(Dst As RECTF, Src As RECTF, Unit As GpUnit) As Long
    GdipBeginContainer mHandle, Dst, Src, Unit, BeginContainer
End Function

Public Function BeginContainer2() As Long
    GdipBeginContainer2 mHandle, BeginContainer2
End Function

Public Sub EndContainer(ByVal state As Long)
    GdipEndContainer mHandle, state
End Sub

Public Function GetNearestColor(ByVal nColor As Long) As Long
    GdipGetNearestColor mHandle, nColor
    GetNearestColor = nColor
End Function

Public Sub Clear(Optional ByVal nColor As Long = &HFFFFFF)
    GdipGraphicsClear mHandle, nColor
End Sub

Public Sub Flush(ByVal Intertion As FlushIntention)
    GdipFlush mHandle, Intertion
End Sub

Public Function IsVisiblePoint(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim Result As Long
    GdipIsVisiblePoint mHandle, X, Y, Result
    IsVisiblePoint = Result
End Function

Public Function IsVisibleRect(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Boolean
    Dim Result As Long
    GdipIsVisibleRect mHandle, X, Y, Width, Height, Result
    IsVisibleRect = Result
End Function

Public Function IsClipEmpty() As Boolean
    Dim Result As Long
    GdipIsClipEmpty mHandle, Result
    IsClipEmpty = Result
End Function

Public Function IsVisibleClipEmpty() As Boolean
    Dim Result As Long
    GdipIsVisibleClipEmpty mHandle, Result
    IsVisibleClipEmpty = Result
End Function

Public Function GetClipBounds() As RECTF
    Dim Rects As RECTF
    GdipGetClipBounds mHandle, Rects
    GetClipBounds = Rects
End Function

Public Function GetVisibleClipBounds() As RECTF
    Dim Rects As RECTF
    GdipGetVisibleClipBounds mHandle, Rects
    GetVisibleClipBounds = Rects
End Function

Public Function GetHalftonePalette() As Long    '???????????
    GetHalftonePalette = GdipCreateHalftonePalette()
End Function

Public Sub DrawArc(Pen As cPen, ByVal X As Single, ByVal Y As Single, ByVal w As Single, ByVal H As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    GdipDrawArc mHandle, Pen.Handle, X, Y, w, H, startAngle, sweepAngle
End Sub

Public Sub DrawBezier(Pen As cPen, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single)
    GdipDrawBezier mHandle, Pen.Handle, x1, y1, x2, y2, x3, y3, x4, y4
End Sub

Public Sub DrawBeziers(Pen As cPen, Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipDrawBeziers mHandle, Pen.Handle, Ps(LBound(Ps)), Count
End Sub

Public Sub DrawCloseCurve(Pen As cPen, Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipDrawClosedCurve mHandle, Pen.Handle, Ps(LBound(Ps)), Count
End Sub

Public Sub DrawCloseCurve2(Pen As cPen, Ps() As POINTF, ByVal tension As Single)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipDrawClosedCurve2 mHandle, Pen.Handle, Ps(LBound(Ps)), Count, tension
End Sub

Public Sub DrawCurve(Pen As cPen, Ps() As POINTF)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipDrawCurve mHandle, Pen.Handle, Ps(LBound(Ps)), Count
End Sub

Public Sub DrawCurve2(Pen As cPen, Ps() As POINTF, ByVal tension As Single)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipDrawCurve2 mHandle, Pen.Handle, Ps(LBound(Ps)), Count, tension
End Sub

Public Sub DrawCurve3(Pen As cPen, Ps() As POINTF, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single)
    Dim Count As Long
    If (Not Ps) = -1 Then Exit Sub
    Count = UBound(Ps) - LBound(Ps) + 1
    GdipDrawCurve3 mHandle, Pen.Handle, Ps(LBound(Ps)), Count, offset, numberOfSegments, tension
End Sub

Public Sub DrawDriverString(ByVal pStr As Long, Font As cFont, Brush As cBrush, Pos As POINTF, Flags As DriverStringOptions, Matrix As cMatrix)
    GdipDrawDriverString mHandle, pStr, -1, Font.Handle, Brush.Handle, Pos, Flags, Matrix.Handle
End Sub

Public Sub DrawString(ByVal sText As String, Font As cFont, Layout As RECTF, strFormat As cStringFormat, Brush As cBrush)
    GdipDrawString mHandle, StrPtr(sText), -1, Font.Handle, Layout, strFormat.Handle, Brush.Handle
End Sub

Public Sub PrintString(ByVal sText As String, Layout As RECTF, Optional nColor As Long = &HFF000000, Optional ByVal nSize As Single = 15, Optional ByVal sFontName As String = "ËÎÌו")
    Dim Fam As Long
    Dim mFont As Long
    Dim mFormat As Long
    Dim Brush As Long
    GdipCreateFontFamilyFromName StrPtr(sFontName), 0, Fam
    GdipCreateFont Fam, nSize, FontStyleRegular, GlobalUnit, mFont
    GdipCreateStringFormat 0, 0, mFormat
    GdipSetStringFormatAlign mFormat, StringAlignmentNear
    GdipCreateSolidFill nColor, Brush
    
    GdipDrawString mHandle, StrPtr(sText), -1, mFont, Layout, mFormat, Brush
    
    GdipDeleteStringFormat mFormat
    GdipDeleteFont mFont
    GdipDeleteFontFamily Fam
End Sub

Public Sub DrawPath(Path As cPath, Pen As cPen)
    GdipDrawPath mHandle, Pen.Handle, Path.Handle
End Sub

Public Function MeasureString(ByVal sText As String, Font As cFont, Layout As RECTF, strFormat As cStringFormat) As RECTF
    GdipMeasureString mHandle, StrPtr(sText), -1, Font.Handle, Layout, strFormat.Handle, MeasureString, 0, 0
End Function

Public Function MeasureDriverString(ByVal pStr As Long, Font As cFont, Brush As cBrush, Pos As POINTF, Flags As DriverStringOptions, Matrix As cMatrix) As RECTF
    GdipMeasureDriverString mHandle, pStr, -1, Font.Handle, Pos, Flags, Matrix.Handle, MeasureDriverString
End Function

Public Function MeasureCharacterRanges(ByVal sText As String, Font As cFont, Layout As RECTF, strFormat As cStringFormat) As cRegion()
    Dim Rgn() As Long
    Dim objRgn() As cRegion
    Dim RgnCount As Long
    Dim i As Integer
    
    RgnCount = strFormat.GetMeasurableCharacterRangeCount
    
    ReDim Rgn(RgnCount - 1) As Long
    ReDim objRgn(RgnCount - 1) As cRegion

    For i = 0 To RgnCount - 1
        GdipCreateRegion Rgn(i)
        Set objRgn(i) = New cRegion
        objRgn(i).Release
        objRgn(i).Handle = Rgn(i)
    Next

    GdipMeasureCharacterRanges mHandle, StrPtr(sText), -1, Font.Handle, Layout, strFormat.Handle, RgnCount, Rgn(0)
    
    MeasureCharacterRanges = objRgn
End Function

Public Sub DrawEllipse(Pen As cPen, ByVal X As Single, ByVal Y As Single, ByVal w As Single, ByVal H As Single)
    GdipDrawEllipse mHandle, Pen.Handle, X, Y, w, H
End Sub

Public Sub DrawLine(Pen As cPen, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    GdipDrawLine mHandle, Pen.Handle, x1, y1, x2, y2
End Sub

Public Sub DrawLines(Pen As cPen, Points() As POINTF)
    Dim Count As Long
    
    If (Not Points) = -1 Then Exit Sub
    
    Count = UBound(Points) - LBound(Points) + 1
    GdipDrawLines mHandle, Pen.Handle, Points(LBound(Points)), Count
End Sub

Public Sub DrawPie(Pen As cPen, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    GdipDrawPie mHandle, Pen.Handle, X, Y, Width, Height, startAngle, sweepAngle
End Sub

Public Sub DrawPolygon(Pen As cPen, Points() As POINTF)
    Dim Count As Long
    
    If (Not Points) = -1 Then Exit Sub
    
    Count = UBound(Points) - LBound(Points) + 1
    GdipDrawPolygon mHandle, Pen.Handle, Points(LBound(Points)), Count
End Sub

Public Sub DrawRectangle(Pen As cPen, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    GdipDrawRectangle mHandle, Pen.Handle, X, Y, Width, Height
End Sub

Public Sub DrawRectangles(Pen As cPen, Rects() As RECTF)
    Dim Count As Long
    
    If (Not Rects) = -1 Then Exit Sub
    
    Count = UBound(Rects) - LBound(Rects) + 1
    GdipDrawRectangles mHandle, Pen.Handle, Rects(LBound(Rects)), Count
End Sub

Public Sub FillCloseCurve(Brush As cBrush, Points() As POINTF)
    Dim Count As Long
    If (Not Points) = -1 Then Exit Sub
    
    Count = UBound(Points) - LBound(Points) + 1
    GdipFillClosedCurve mHandle, Brush.Handle, Points(LBound(Points)), Count
End Sub

Public Sub FillCloseCurve2(Brush As cBrush, Points() As POINTF, ByVal tension As Single, FillMd As FillMode)
    Dim Count As Long
    If (Not Points) = -1 Then Exit Sub
    
    Count = UBound(Points) - LBound(Points) + 1
    GdipFillClosedCurve2 mHandle, Brush.Handle, Points(LBound(Points)), Count, tension, FillMd
End Sub

Public Sub FillPolygon(Brush As cBrush, Points() As POINTF, FillMd As FillMode)
    Dim Count As Long
    If (Not Points) = -1 Then Exit Sub
    
    Count = UBound(Points) - LBound(Points) + 1
    GdipFillPolygon mHandle, Brush.Handle, Points(LBound(Points)), Count, FillMd
End Sub

Public Sub FillPolygon2(Brush As cBrush, Points() As POINTF)
    Dim Count As Long
    If (Not Points) = -1 Then Exit Sub
    
    Count = UBound(Points) - LBound(Points) + 1
    GdipFillPolygon2 mHandle, Brush.Handle, Points(LBound(Points)), Count
End Sub

Public Sub FillRectangle(Brush As cBrush, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    GdipFillRectangle mHandle, Brush.Handle, X, Y, Width, Height
End Sub

Public Sub FillRectangles(Brush As cBrush, Rects() As RECTF)
    Dim Count As Long
    If (Not Rects) = -1 Then Exit Sub
    
    Count = UBound(Rects) - LBound(Rects) + 1
    GdipFillRectangles mHandle, Brush.Handle, Rects(LBound(Rects)), Count
End Sub

Public Sub FillEllipse(Brush As cBrush, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    GdipFillEllipse mHandle, Brush.Handle, X, Y, Width, Height
End Sub

Public Sub FillPie(Brush As cBrush, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    GdipFillPie mHandle, Brush.Handle, X, Y, Width, Height, startAngle, sweepAngle
End Sub

Public Sub FillPath(Brush As cBrush, Path As cPath)
    GdipFillPath mHandle, Brush.Handle, Path.Handle
End Sub

Public Sub FillRegion(Brush As cBrush, Rgn As cRegion)
    GdipFillRegion mHandle, Brush.Handle, Rgn.Handle
End Sub

Public Sub DrawChachedBitmap(Cached As cCachedBitmap, ByVal X As Long, ByVal Y As Long)
    GdipDrawCachedBitmap mHandle, Cached.Handle, X, Y
End Sub

Public Sub DrawImage(Img As cImage, ByVal X As Single, ByVal Y As Single)
     GdipDrawImage mHandle, Img.Handle, X, Y
End Sub

Public Sub DrawImageRect(Img As cImage, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    p_DrawImageRect Img.Handle, X, Y, Width, Height
End Sub

Public Sub DrawImageRectRect(Img As cImage, dstRect As RECTF, srcRect As RECTF, imgAttr As cImageAttributes)
    p_DrawImageRectRect Img.Handle, dstRect, srcRect, imgAttr
End Sub

Public Sub DrawImagePointRect(Img As cImage, ByVal X As Single, ByVal Y As Single, srcRect As RECTF)
    p_DrawImagePointRect Img.Handle, X, Y, srcRect
End Sub

Public Sub DrawImagePointsRect(Img As cImage, pts() As POINTF, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single)
    p_DrawImagePointsRect Img.Handle, pts, SrcX, SrcY, srcWidth, srcHeight
End Sub

Public Sub DrawImagePoints(Img As cImage, pts() As POINTF)
    p_DrawImagePoints Img.Handle, pts
End Sub

Public Sub DrawBitmap(Img As cBitmap, ByVal X As Single, ByVal Y As Single)
     GdipDrawImage mHandle, Img.Handle, X, Y
End Sub

Public Sub DrawBitmapRect(Img As cBitmap, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    p_DrawImageRect Img.Handle, X, Y, Width, Height
End Sub

Public Sub DrawBitmapRectRect(Img As cBitmap, dstRect As RECTF, srcRect As RECTF, imgAttr As cImageAttributes)
    p_DrawImageRectRect Img.Handle, dstRect, srcRect, imgAttr
End Sub

Public Sub DrawBitmapPointRect(Img As cBitmap, ByVal X As Single, ByVal Y As Single, srcRect As RECTF)
    p_DrawImagePointRect Img.Handle, X, Y, srcRect
End Sub

Public Sub DrawBitmapPointsRect(Img As cBitmap, pts() As POINTF, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single)
    p_DrawImagePointsRect Img.Handle, pts, SrcX, SrcY, srcWidth, srcHeight
End Sub

Public Sub DrawBitmapPoints(Img As cBitmap, pts() As POINTF)
    p_DrawImagePoints Img.Handle, pts
End Sub

Private Sub p_DrawImageRect(Img As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    GdipDrawImageRect mHandle, Img, X, Y, Width, Height
End Sub

Private Sub p_DrawImageRectRect(Img As Long, dstRect As RECTF, srcRect As RECTF, imgAttr As cImageAttributes)
    Dim H As Long
    If Not imgAttr Is Nothing Then H = imgAttr.Handle
    GdipDrawImageRectRect mHandle, Img, dstRect.Left, dstRect.Top, dstRect.Right, dstRect.Bottom, srcRect.Left, srcRect.Top, srcRect.Right, srcRect.Bottom, GlobalUnit, H, 0, 0
End Sub

Private Sub p_DrawImagePointRect(Img As Long, ByVal X As Single, ByVal Y As Single, srcRect As RECTF)
    GdipDrawImagePointRect mHandle, Img, X, Y, srcRect.Left, srcRect.Top, srcRect.Right, srcRect.Bottom, GlobalUnit
End Sub

Private Sub p_DrawImagePointsRect(Img As Long, pts() As POINTF, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single)
    Dim Count As Long
    If (Not pts) = -1 Then Exit Sub
    Count = UBound(pts) - LBound(pts) + 1
    GdipDrawImagePointsRect mHandle, Img, pts(LBound(pts)), Count, SrcX, SrcY, srcWidth, srcHeight, GlobalUnit
End Sub

Private Sub p_DrawImagePoints(Img As Long, pts() As POINTF)
    Dim Count As Long
    If (Not pts) = -1 Then Exit Sub
    Count = UBound(pts) - LBound(pts) + 1
    GdipDrawImagePoints mHandle, Img, pts(LBound(pts)), Count
End Sub
