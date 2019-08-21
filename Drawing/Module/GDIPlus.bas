Attribute VB_Name = "GDIPlus"
Option Explicit

'#Const GdipVersion = 1.1

Public Const ImageEncoderSuffix       As String = "-1A04-11D3-9A73-0000F81EF32E}"
Public Const ImageEncoderBMP          As String = "{557CF400" & ImageEncoderSuffix
Public Const ImageEncoderJPG          As String = "{557CF401" & ImageEncoderSuffix
Public Const ImageEncoderGIF          As String = "{557CF402" & ImageEncoderSuffix
Public Const ImageEncoderEMF          As String = "{557CF403" & ImageEncoderSuffix
Public Const ImageEncoderWMF          As String = "{557CF404" & ImageEncoderSuffix
Public Const ImageEncoderTIF          As String = "{557CF405" & ImageEncoderSuffix
Public Const ImageEncoderPNG          As String = "{557CF406" & ImageEncoderSuffix
Public Const ImageEncoderICO          As String = "{557CF407" & ImageEncoderSuffix
Public Const EncoderCompression       As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Public Const EncoderColorDepth        As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Public Const EncoderScanMethod        As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Public Const EncoderVersion           As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Public Const EncoderRenderMethod      As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Public Const EncoderQuality           As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Public Const EncoderTransformation    As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Public Const EncoderLuminanceTable    As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Public Const EncoderChrominanceTable  As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Public Const EncoderSaveFlag          As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
Public Const EncoderColorSpace        As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
Public Const EncoderImageItems        As String = "{63875E13-1F1D-45AB-9195-A29B6066A650}"
Public Const EncoderSaveAsCMYK        As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"

Public Declare Function GdipGetDC Lib "gdiplus" (ByVal Graphics As Long, hdc As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal Graphics As Long, ByVal hdc As Long) As GpStatus
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, Graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal Hwnd As Long, Graphics As Long) As GpStatus
Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, Graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Public Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal Graphics As Long, ByVal lColor As Long) As GpStatus
Public Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal CompositingMd As CompositingMode) As GpStatus
Public Declare Function GdipGetCompositingMode Lib "gdiplus" (ByVal Graphics As Long, CompositingMd As CompositingMode) As GpStatus
Public Declare Function GdipSetRenderingOrigin Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long) As GpStatus
Public Declare Function GdipGetRenderingOrigin Lib "gdiplus" (ByVal Graphics As Long, X As Long, Y As Long) As GpStatus
Public Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal Graphics As Long, ByVal CompositingQlty As CompositingQuality) As GpStatus
Public Declare Function GdipGetCompositingQuality Lib "gdiplus" (ByVal Graphics As Long, CompositingQlty As CompositingQuality) As GpStatus
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMd As SmoothingMode) As GpStatus
Public Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, SmoothingMd As SmoothingMode) As GpStatus
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixOffsetMode As PixelOffsetMode) As GpStatus
Public Declare Function GdipGetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, PixOffsetMode As PixelOffsetMode) As GpStatus
Public Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal Graphics As Long, ByVal Mode As TextRenderingHint) As GpStatus
Public Declare Function GdipGetTextRenderingHint Lib "gdiplus" (ByVal Graphics As Long, Mode As TextRenderingHint) As GpStatus
Public Declare Function GdipSetTextContrast Lib "gdiplus" (ByVal Graphics As Long, ByVal Contrast As Long) As GpStatus
Public Declare Function GdipGetTextContrast Lib "gdiplus" (ByVal Graphics As Long, Contrast As Long) As GpStatus
Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal interpolation As InterpolationMode) As GpStatus
Public Declare Function GdipGetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, interpolation As InterpolationMode) As GpStatus
Public Declare Function GdipSetWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetWorldTransform Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Public Declare Function GdipMultiplyWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal sx As Single, ByVal sy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipGetWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetPageTransform Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Public Declare Function GdipGetPageUnit Lib "gdiplus" (ByVal Graphics As Long, Unit As GpUnit) As GpStatus
Public Declare Function GdipGetPageScale Lib "gdiplus" (ByVal Graphics As Long, sScale As Single) As GpStatus
Public Declare Function GdipSetPageUnit Lib "gdiplus" (ByVal Graphics As Long, ByVal Unit As GpUnit) As GpStatus
Public Declare Function GdipSetPageScale Lib "gdiplus" (ByVal Graphics As Long, ByVal sScale As Single) As GpStatus
Public Declare Function GdipGetDpiX Lib "gdiplus" (ByVal Graphics As Long, DPI As Single) As GpStatus
Public Declare Function GdipGetDpiY Lib "gdiplus" (ByVal Graphics As Long, DPI As Single) As GpStatus
Public Declare Function GdipTransformPoints Lib "gdiplus" (ByVal Graphics As Long, ByVal destSpace As CoordinateSpace, ByVal srcSpace As CoordinateSpace, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformPointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal destSpace As CoordinateSpace, ByVal srcSpace As CoordinateSpace, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformPoints_ Lib "gdiplus" Alias "GdipTransformPoints" (ByVal Graphics As Long, ByVal destSpace As CoordinateSpace, ByVal srcSpace As CoordinateSpace, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformPointsI_ Lib "gdiplus" Alias "GdipTransformPointsI" (ByVal Graphics As Long, ByVal destSpace As CoordinateSpace, ByVal srcSpace As CoordinateSpace, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetNearestColor Lib "gdiplus" (ByVal Graphics As Long, ARGB As Long) As GpStatus
Public Declare Function GdipCreateHalftonePalette Lib "gdiplus" () As Long
Public Declare Function GdipSetClipGraphics Lib "gdiplus" (ByVal Graphics As Long, ByVal srcgraphics As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRect Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Path As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal Graphics As Long, ByVal region As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipHrgn Lib "gdiplus" (ByVal Graphics As Long, ByVal hRgn As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipResetClip Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Public Declare Function GdipTranslateClip Lib "gdiplus" (ByVal Graphics As Long, ByVal dx As Single, ByVal dy As Single) As GpStatus
Public Declare Function GdipTranslateClipI Lib "gdiplus" (ByVal Graphics As Long, ByVal dx As Long, ByVal dy As Long) As GpStatus
Public Declare Function GdipGetClip Lib "gdiplus" (ByVal Graphics As Long, ByVal region As Long) As GpStatus
Public Declare Function GdipGetClipBounds Lib "gdiplus" (ByVal Graphics As Long, RECT As RECTF) As GpStatus
Public Declare Function GdipGetClipBoundsI Lib "gdiplus" (ByVal Graphics As Long, RECT As RECTL) As GpStatus
Public Declare Function GdipIsClipEmpty Lib "gdiplus" (ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipGetVisibleClipBounds Lib "gdiplus" (ByVal Graphics As Long, RECT As RECTF) As GpStatus
Public Declare Function GdipGetVisibleClipBoundsI Lib "gdiplus" (ByVal Graphics As Long, RECT As RECTL) As GpStatus
Public Declare Function GdipIsVisibleClipEmpty Lib "gdiplus" (ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisiblePoint Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Single, ByVal Y As Single, Result As Long) As GpStatus
Public Declare Function GdipIsVisiblePointI Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisibleRect Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Result As Long) As GpStatus
Public Declare Function GdipIsVisibleRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Result As Long) As GpStatus
Public Declare Function GdipSaveGraphics Lib "gdiplus" (ByVal Graphics As Long, state As Long) As GpStatus
Public Declare Function GdipRestoreGraphics Lib "gdiplus" (ByVal Graphics As Long, ByVal state As Long) As GpStatus
Public Declare Function GdipBeginContainer Lib "gdiplus" (ByVal Graphics As Long, dstRect As RECTF, srcRect As RECTF, ByVal Unit As GpUnit, state As Long) As GpStatus
Public Declare Function GdipBeginContainerI Lib "gdiplus" (ByVal Graphics As Long, dstRect As RECTL, srcRect As RECTL, ByVal Unit As GpUnit, state As Long) As GpStatus
Public Declare Function GdipBeginContainer2 Lib "gdiplus" (ByVal Graphics As Long, state As Long) As GpStatus
Public Declare Function GdipEndContainer Lib "gdiplus" (ByVal Graphics As Long, ByVal state As Long) As GpStatus
Public Declare Function GdipDrawLine Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GpStatus
Public Declare Function GdipDrawLineI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GpStatus
Public Declare Function GdipDrawLines Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLines_ Lib "gdiplus" Alias "GdipDrawLines" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLinesI_ Lib "gdiplus" Alias "GdipDrawLinesI" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawArc Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawArcI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawBezier Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As GpStatus
Public Declare Function GdipDrawBezierI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As GpStatus
Public Declare Function GdipDrawBeziers Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziersI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziers_ Lib "gdiplus" Alias "GdipDrawBeziers" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziersI_ Lib "gdiplus" Alias "GdipDrawBeziersI" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawRectangles Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Rects As RECTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawRectanglesI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Rects As RECTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillRectangle Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillRectangles Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Rects As RECTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillRectanglesI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Rects As RECTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillEllipse Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawPie Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawPieI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipFillPie Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipFillPieI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawPolygon Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillPolygon Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTF, ByVal Count As Long, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTL, ByVal Count As Long, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygon2 Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillPolygon2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygon_ Lib "gdiplus" Alias "GdipDrawPolygon" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygonI_ Lib "gdiplus" Alias "GdipDrawPolygonI" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillPolygon_ Lib "gdiplus" Alias "GdipFillPolygon" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygonI_ Lib "gdiplus" Alias "GdipFillPolygonI" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygon2_ Lib "gdiplus" Alias "GdipFillPolygon2" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillPolygon2I_ Lib "gdiplus" Alias "GdipFillPolygon2I" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipFillPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipDrawCurve Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurveI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurve2 Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTL, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3 Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal Count As Long, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3I Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTL, ByVal Count As Long, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurveI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurve2 Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTL, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipFillClosedCurve Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurveI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurve2 Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTF, ByVal Count As Long, ByVal tension As Single, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTL, ByVal Count As Long, ByVal tension As Single, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipDrawCurve_ Lib "gdiplus" Alias "GdipDrawCurve" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurveI_ Lib "gdiplus" Alias "GdipDrawCurveI" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurve2_ Lib "gdiplus" Alias "GdipDrawCurve2" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve2I_ Lib "gdiplus" Alias "GdipDrawCurve2I" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3_ Lib "gdiplus" Alias "GdipDrawCurve3" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3I_ Lib "gdiplus" Alias "GdipDrawCurve3I" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve_ Lib "gdiplus" Alias "GdipDrawClosedCurve" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurveI_ Lib "gdiplus" Alias "GdipDrawClosedCurveI" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurve2_ Lib "gdiplus" Alias "GdipDrawClosedCurve2" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve2I_ Lib "gdiplus" Alias "GdipDrawClosedCurve2I" (ByVal Graphics As Long, ByVal Pen As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipFillClosedCurve_ Lib "gdiplus" Alias "GdipFillClosedCurve" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurveI_ Lib "gdiplus" Alias "GdipFillClosedCurveI" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurve2_ Lib "gdiplus" Alias "GdipFillClosedCurve2" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long, ByVal tension As Single, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillClosedCurve2I_ Lib "gdiplus" Alias "GdipFillClosedCurve2I" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long, ByVal tension As Single, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillRegion Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal region As Long) As GpStatus
Public Declare Function GdipDrawImage Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single) As GpStatus
Public Declare Function GdipDrawImageI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImagePoints Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, dstpoints As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, dstpoints As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePoints_ Lib "gdiplus" Alias "GdipDrawImagePoints" (ByVal Graphics As Long, ByVal Image As Long, dstpoints As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI_ Lib "gdiplus" Alias "GdipDrawImagePointsI" (ByVal Graphics As Long, ByVal Image As Long, dstpoints As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImagePointRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImagePointsRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, Points As POINTF, ByVal Count As Long, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
Public Declare Function GdipDrawImagePointsRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, Points As POINTL, ByVal Count As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRect_ Lib "gdiplus" Alias "GdipDrawImagePointsRect" (ByVal Graphics As Long, ByVal Image As Long, Points As Any, ByVal Count As Long, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRectI_ Lib "gdiplus" Alias "GdipDrawImagePointsRectI" (ByVal Graphics As Long, ByVal Image As Long, Points As Any, ByVal Count As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Public Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal numDecoders As Long, ByVal Size As Long, decoders As Any) As GpStatus
Public Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GpStatus
Public Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, encoders As Any) As GpStatus
Public Declare Function GdipComment Lib "gdiplus" (ByVal Graphics As Long, ByVal sizeData As Long, data As Any) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromFileICM Lib "gdiplus" (ByVal FileName As Long, Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As Any, Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStreamICM Lib "gdiplus" (ByVal stream As Any, Image As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As Any, encoderParams As Any) As GpStatus
Public Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal stream As Any, clsidEncoder As CLSID, encoderParams As Any) As GpStatus
Public Declare Function GdipSaveAdd Lib "gdiplus" (ByVal Image As Long, encoderParams As EncoderParameters) As GpStatus
Public Declare Function GdipSaveAddImage Lib "gdiplus" (ByVal Image As Long, ByVal newImage As Long, encoderParams As EncoderParameters) As GpStatus
Public Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal Image As Long, srcRect As RECTF, srcUnit As GpUnit) As GpStatus
Public Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, Width As Single, Height As Single) As GpStatus
Public Declare Function GdipGetImageType Lib "gdiplus" (ByVal Image As Long, itype As Image_Type) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Public Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal Image As Long, resolution As Single) As GpStatus
Public Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal Image As Long, resolution As Single) As GpStatus
Public Declare Function GdipGetImageFlags Lib "gdiplus" (ByVal Image As Long, Flags As Long) As GpStatus
Public Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal Image As Long, Format As CLSID) As GpStatus
Public Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, PixelFormat As GpPixelFormat) As GpStatus
Public Declare Function GdipGetImageThumbnail Lib "gdiplus" (ByVal Image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Public Declare Function GdipGetEncoderParameterListSize Lib "gdiplus" (ByVal Image As Long, clsidEncoder As CLSID, Size As Long) As GpStatus
Public Declare Function GdipGetEncoderParameterList Lib "gdiplus" (ByVal Image As Long, clsidEncoder As CLSID, ByVal Size As Long, Buffer As EncoderParameters) As GpStatus
Public Declare Function GdipImageGetFrameDimensionsCount Lib "gdiplus" (ByVal Image As Long, Count As Long) As GpStatus
Public Declare Function GdipImageGetFrameDimensionsList Lib "gdiplus" (ByVal Image As Long, dimensionIDs As CLSID, ByVal Count As Long) As GpStatus
Public Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal Image As Long, dimensionID As CLSID, Count As Long) As GpStatus
Public Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal Image As Long, dimensionID As CLSID, ByVal frameIndex As Long) As GpStatus
Public Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As RotateFlipType) As GpStatus
Public Declare Function GdipGetImagePalette Lib "gdiplus" (ByVal Image As Long, Palette As ColorPalette, ByVal Size As Long) As GpStatus
Public Declare Function GdipSetImagePalette Lib "gdiplus" (ByVal Image As Long, Palette As ColorPalette) As GpStatus
Public Declare Function GdipGetImagePaletteSize Lib "gdiplus" (ByVal Image As Long, Size As Long) As GpStatus
Public Declare Function GdipGetPropertyCount Lib "gdiplus" (ByVal Image As Long, numOfProperty As Long) As GpStatus
Public Declare Function GdipGetPropertyIdList Lib "gdiplus" (ByVal Image As Long, ByVal numOfProperty As Long, List As Long) As GpStatus
Public Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, Size As Long) As GpStatus
Public Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, ByVal propSize As Long, Buffer As PropertyItem) As GpStatus
Public Declare Function GdipGetPropertySize Lib "gdiplus" (ByVal Image As Long, totalBufferSize As Long, numProperties As Long) As GpStatus
Public Declare Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal Image As Long, ByVal totalBufferSize As Long, ByVal numProperties As Long, allItems As PropertyItem) As GpStatus
Public Declare Function GdipRemovePropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long) As GpStatus
Public Declare Function GdipSetPropertyItem Lib "gdiplus" (ByVal Image As Long, Item As PropertyItem) As GpStatus
Public Declare Function GdipImageForceValidation Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal Unit As GpUnit, Pen As Long) As GpStatus
Public Declare Function GdipCreatePen2 Lib "gdiplus" (ByVal Brush As Long, ByVal Width As Single, ByVal Unit As GpUnit, Pen As Long) As GpStatus
Public Declare Function GdipClonePen Lib "gdiplus" (ByVal Pen As Long, clonepen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal Pen As Long) As GpStatus
Public Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal Pen As Long, ByVal Width As Single) As GpStatus
Public Declare Function GdipGetPenWidth Lib "gdiplus" (ByVal Pen As Long, Width As Single) As GpStatus
Public Declare Function GdipSetPenUnit Lib "gdiplus" (ByVal Pen As Long, ByVal Unit As GpUnit) As GpStatus
Public Declare Function GdipGetPenUnit Lib "gdiplus" (ByVal Pen As Long, Unit As GpUnit) As GpStatus
Public Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal Pen As Long, ByVal StartCap As LineCap, ByVal EndCap As LineCap, ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal Pen As Long, ByVal StartCap As LineCap) As GpStatus
Public Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal Pen As Long, ByVal EndCap As LineCap) As GpStatus
Public Declare Function GdipSetPenDashCap Lib "gdiplus" Alias "GdipSetPenDashCap197819" (ByVal Pen As Long, ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipGetPenStartCap Lib "gdiplus" (ByVal Pen As Long, StartCap As LineCap) As GpStatus
Public Declare Function GdipGetPenEndCap Lib "gdiplus" (ByVal Pen As Long, EndCap As LineCap) As GpStatus
Public Declare Function GdipGetPenDashCap Lib "gdiplus" Alias "GdipGetPenDashCap197819" (ByVal Pen As Long, dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal Pen As Long, ByVal lnJoin As GpLineJoin) As GpStatus
Public Declare Function GdipGetPenLineJoin Lib "gdiplus" (ByVal Pen As Long, lnJoin As GpLineJoin) As GpStatus
Public Declare Function GdipSetPenCustomStartCap Lib "gdiplus" (ByVal Pen As Long, ByVal customCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomStartCap Lib "gdiplus" (ByVal Pen As Long, customCap As Long) As GpStatus
Public Declare Function GdipSetPenCustomEndCap Lib "gdiplus" (ByVal Pen As Long, ByVal customCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomEndCap Lib "gdiplus" (ByVal Pen As Long, customCap As Long) As GpStatus
Public Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal Pen As Long, ByVal MiterLimit As Single) As GpStatus
Public Declare Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal Pen As Long, MiterLimit As Single) As GpStatus
Public Declare Function GdipSetPenMode Lib "gdiplus" (ByVal Pen As Long, ByVal PenMode As PenAlignment) As GpStatus
Public Declare Function GdipGetPenMode Lib "gdiplus" (ByVal Pen As Long, PenMode As PenAlignment) As GpStatus
Public Declare Function GdipSetPenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetPenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetPenTransform Lib "gdiplus" (ByVal Pen As Long) As GpStatus
Public Declare Function GdipMultiplyPenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal sx As Single, ByVal sy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipSetPenColor Lib "gdiplus" (ByVal Pen As Long, ByVal ARGB As Long) As GpStatus
Public Declare Function GdipGetPenColor Lib "gdiplus" (ByVal Pen As Long, ARGB As Long) As GpStatus
Public Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal Pen As Long, ByVal Brush As Long) As GpStatus
Public Declare Function GdipGetPenBrushFill Lib "gdiplus" (ByVal Pen As Long, Brush As Long) As GpStatus
Public Declare Function GdipGetPenFillType Lib "gdiplus" (ByVal Pen As Long, pType As PenType) As GpStatus
Public Declare Function GdipGetPenDashStyle Lib "gdiplus" (ByVal Pen As Long, dStyle As DashStyle) As GpStatus
Public Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal Pen As Long, ByVal dStyle As DashStyle) As GpStatus
Public Declare Function GdipGetPenDashOffset Lib "gdiplus" (ByVal Pen As Long, offset As Single) As GpStatus
Public Declare Function GdipSetPenDashOffset Lib "gdiplus" (ByVal Pen As Long, ByVal offset As Single) As GpStatus
Public Declare Function GdipGetPenDashCount Lib "gdiplus" (ByVal Pen As Long, Count As Long) As GpStatus
Public Declare Function GdipSetPenDashArray Lib "gdiplus" (ByVal Pen As Long, dash As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenDashArray Lib "gdiplus" (ByVal Pen As Long, dash As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundCount Lib "gdiplus" (ByVal Pen As Long, Count As Long) As GpStatus
Public Declare Function GdipSetPenCompoundArray Lib "gdiplus" (ByVal Pen As Long, dash As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundArray Lib "gdiplus" (ByVal Pen As Long, dash As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipCreateCustomLineCap Lib "gdiplus" (ByVal fillPath As Long, ByVal strokePath As Long, ByVal BaseCap As LineCap, ByVal BaseInset As Single, customCap As Long) As GpStatus
Public Declare Function GdipDeleteCustomLineCap Lib "gdiplus" (ByVal customCap As Long) As GpStatus
Public Declare Function GdipCloneCustomLineCap Lib "gdiplus" (ByVal customCap As Long, clonedCap As Long) As GpStatus
Public Declare Function GdipGetCustomLineCapType Lib "gdiplus" (ByVal customCap As Long, capType As CustomLineCapType) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeCaps Lib "gdiplus" (ByVal customCap As Long, ByVal StartCap As LineCap, ByVal EndCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeCaps Lib "gdiplus" (ByVal customCap As Long, StartCap As LineCap, EndCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeJoin Lib "gdiplus" (ByVal customCap As Long, ByVal lnJoin As GpLineJoin) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeJoin Lib "gdiplus" (ByVal customCap As Long, lnJoin As GpLineJoin) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseCap Lib "gdiplus" (ByVal customCap As Long, ByVal BaseCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseCap Lib "gdiplus" (ByVal customCap As Long, BaseCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseInset Lib "gdiplus" (ByVal customCap As Long, ByVal inset As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseInset Lib "gdiplus" (ByVal customCap As Long, inset As Single) As GpStatus
Public Declare Function GdipSetCustomLineCapWidthScale Lib "gdiplus" (ByVal customCap As Long, ByVal WidthScale As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapWidthScale Lib "gdiplus" (ByVal customCap As Long, WidthScale As Single) As GpStatus
Public Declare Function GdipCreateAdjustableArrowCap Lib "gdiplus" (ByVal Height As Single, ByVal Width As Single, ByVal isFilled As Long, Cap As Long) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapHeight Lib "gdiplus" (ByVal Cap As Long, ByVal Height As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapHeight Lib "gdiplus" (ByVal Cap As Long, Height As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapWidth Lib "gdiplus" (ByVal Cap As Long, ByVal Width As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapWidth Lib "gdiplus" (ByVal Cap As Long, Width As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapMiddleInset Lib "gdiplus" (ByVal Cap As Long, ByVal middleInset As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapMiddleInset Lib "gdiplus" (ByVal Cap As Long, middleInset As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapFillState Lib "gdiplus" (ByVal Cap As Long, ByVal bFillState As Long) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapFillState Lib "gdiplus" (ByVal Cap As Long, bFillState As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFileICM Lib "gdiplus" (ByVal FileName As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal stream As Any, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromStreamICM Lib "gdiplus" (ByVal stream As Any, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As GpPixelFormat, Scan0 As Any, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Graphics As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, ByVal gdiBitmapData As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, hbmReturn As Long, ByVal background As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hIcon As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, hbmReturn As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromResource Lib "gdiplus" (ByVal hInstance As Long, ByVal lpBitmapName As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCloneBitmapArea Lib "gdiplus" (ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal PixelFormat As GpPixelFormat, ByVal srcBitmap As Long, dstBitmap As Long) As GpStatus
Public Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As GpPixelFormat, ByVal srcBitmap As Long, dstBitmap As Long) As GpStatus
Public Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal Bitmap As Long, RECT As RECTL, ByVal Flags As ImageLockMode, ByVal PixelFormat As GpPixelFormat, lockedBitmapData As BitmapData) As GpStatus
Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal Bitmap As Long, lockedBitmapData As BitmapData) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal Bitmap As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal Bitmap As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As GpStatus
Public Declare Function GdipBitmapSetResolution Lib "gdiplus" (ByVal Bitmap As Long, ByVal xDpi As Single, ByVal yDpi As Single) As GpStatus
Public Declare Function GdipCreateCachedBitmap Lib "gdiplus" (ByVal Bitmap As Long, ByVal Graphics As Long, cachedBitmap As Long) As GpStatus
Public Declare Function GdipDeleteCachedBitmap Lib "gdiplus" (ByVal cachedBitmap As Long) As GpStatus
Public Declare Function GdipDrawCachedBitmap Lib "gdiplus" (ByVal Graphics As Long, ByVal cachedBitmap As Long, ByVal X As Long, ByVal Y As Long) As GpStatus
Public Declare Function GdipCloneBrush Lib "gdiplus" (ByVal Brush As Long, cloneBrush As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipGetBrushType Lib "gdiplus" (ByVal Brush As Long, brshType As BrushType) As GpStatus
Public Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal Style As HatchStyle, ByVal forecolr As Long, ByVal backcolr As Long, Brush As Long) As GpStatus
Public Declare Function GdipGetHatchStyle Lib "gdiplus" (ByVal Brush As Long, Style As HatchStyle) As GpStatus
Public Declare Function GdipGetHatchForegroundColor Lib "gdiplus" (ByVal Brush As Long, forecolr As Long) As GpStatus
Public Declare Function GdipGetHatchBackgroundColor Lib "gdiplus" (ByVal Brush As Long, backcolr As Long) As GpStatus
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, Brush As Long) As GpStatus
Public Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, ByVal ARGB As Long) As GpStatus
Public Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, ARGB As Long) As GpStatus
Public Declare Function GdipCreateLineBrush Lib "gdiplus" (Point1 As POINTF, Point2 As POINTF, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushI Lib "gdiplus" (Point1 As POINTL, Point2 As POINTL, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" (RECT As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal Mode As LinearGradientMode, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (RECT As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal Mode As LinearGradientMode, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (RECT As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal Angle As Single, ByVal isAngleScalable As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "gdiplus" (RECT As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal Angle As Single, ByVal isAngleScalable As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipSetLineColors Lib "gdiplus" (ByVal Brush As Long, ByVal color1 As Long, ByVal color2 As Long) As GpStatus
Public Declare Function GdipGetLineColors Lib "gdiplus" (ByVal Brush As Long, lColors As Long) As GpStatus
Public Declare Function GdipGetLineRect Lib "gdiplus" (ByVal Brush As Long, RECT As RECTF) As GpStatus
Public Declare Function GdipGetLineRectI Lib "gdiplus" (ByVal Brush As Long, RECT As RECTL) As GpStatus
Public Declare Function GdipSetLineGammaCorrection Lib "gdiplus" (ByVal Brush As Long, ByVal useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineGammaCorrection Lib "gdiplus" (ByVal Brush As Long, useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineBlendCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetLineBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLineBlend_ Lib "gdiplus" Alias "GdipGetLineBlend" (ByVal Brush As Long, Blend As Any, positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineBlend_ Lib "gdiplus" Alias "GdipSetLineBlend" (ByVal Brush As Long, Blend As Any, positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlendCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlend_ Lib "gdiplus" Alias "GdipGetLinePresetBlend" (ByVal Brush As Long, Blend As Any, positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLinePresetBlend_ Lib "gdiplus" Alias "GdipSetLinePresetBlend" (ByVal Brush As Long, Blend As Any, positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineSigmaBlend Lib "gdiplus" (ByVal Brush As Long, ByVal Focus As Single, ByVal theScale As Single) As GpStatus
Public Declare Function GdipSetLineLinearBlend Lib "gdiplus" (ByVal Brush As Long, ByVal Focus As Single, ByVal theScale As Single) As GpStatus
Public Declare Function GdipSetLineWrapMode Lib "gdiplus" (ByVal Brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineWrapMode Lib "gdiplus" (ByVal Brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineTransform Lib "gdiplus" (ByVal Brush As Long, Matrix As Long) As GpStatus
Public Declare Function GdipSetLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetLineTransform Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipMultiplyLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipCreateTexture Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As WrapMode, texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2 Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As WrapMode, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIA Lib "gdiplus" (ByVal Image As Long, ByVal imageAttributes As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2I Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As WrapMode, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIAI Lib "gdiplus" (ByVal Image As Long, ByVal imageAttributes As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, texture As Long) As GpStatus
Public Declare Function GdipGetTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipSetTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetTextureTransform Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipTranslateTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipMultiplyTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal Brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal Brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureImage Lib "gdiplus" (ByVal Brush As Long, Image As Long) As GpStatus
Public Declare Function GdipCreatePathGradient Lib "gdiplus" (Points As POINTF, ByVal Count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientI Lib "gdiplus" (Points As POINTL, ByVal Count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradient_ Lib "gdiplus" Alias "GdipCreatePathGradient" (Points As Any, ByVal Count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientI_ Lib "gdiplus" Alias "GdipCreatePathGradientI" (Points As Any, ByVal Count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal Path As Long, polyGradient As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterColor Lib "gdiplus" (ByVal Brush As Long, lColors As Long) As GpStatus
Public Declare Function GdipSetPathGradientCenterColor Lib "gdiplus" (ByVal Brush As Long, ByVal lColors As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorsWithCount Lib "gdiplus" (ByVal Brush As Long, ARGB As Long, Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "gdiplus" (ByVal Brush As Long, ARGB As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPath Lib "gdiplus" (ByVal Brush As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipSetPathGradientPath Lib "gdiplus" (ByVal Brush As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterPoint Lib "gdiplus" (ByVal Brush As Long, Points As POINTF) As GpStatus
Public Declare Function GdipGetPathGradientCenterPointI Lib "gdiplus" (ByVal Brush As Long, Points As POINTL) As GpStatus
Public Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal Brush As Long, Points As POINTF) As GpStatus
Public Declare Function GdipSetPathGradientCenterPointI Lib "gdiplus" (ByVal Brush As Long, Points As POINTL) As GpStatus
Public Declare Function GdipGetPathGradientRect Lib "gdiplus" (ByVal Brush As Long, RECT As RECTF) As GpStatus
Public Declare Function GdipGetPathGradientRectI Lib "gdiplus" (ByVal Brush As Long, RECT As RECTL) As GpStatus
Public Declare Function GdipGetPathGradientPointCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientGammaCorrection Lib "gdiplus" (ByVal Brush As Long, ByVal useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientGammaCorrection Lib "gdiplus" (ByVal Brush As Long, useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlendCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlend_ Lib "gdiplus" Alias "GdipGetPathGradientBlend" (ByVal Brush As Long, Blend As Any, positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientBlend_ Lib "gdiplus" Alias "GdipSetPathGradientBlend" (ByVal Brush As Long, Blend As Any, positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlendCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlend_ Lib "gdiplus" Alias "GdipGetPathGradientPresetBlend" (ByVal Brush As Long, Blend As Any, positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientPresetBlend_ Lib "gdiplus" Alias "GdipSetPathGradientPresetBlend" (ByVal Brush As Long, Blend As Any, positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSigmaBlend Lib "gdiplus" (ByVal Brush As Long, ByVal Focus As Single, ByVal sScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientLinearBlend Lib "gdiplus" (ByVal Brush As Long, ByVal Focus As Single, ByVal sScale As Single) As GpStatus
Public Declare Function GdipGetPathGradientWrapMode Lib "gdiplus" (ByVal Brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal Brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetPathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipSetPathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetPathGradientTransform Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipMultiplyPathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipGetPathGradientFocusScales Lib "gdiplus" (ByVal Brush As Long, xScale As Single, yScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientFocusScales Lib "gdiplus" (ByVal Brush As Long, ByVal xScale As Single, ByVal yScale As Single) As GpStatus
Public Declare Function GdipCreatePath Lib "gdiplus" (ByVal BrushMode As FillMode, Path As Long) As GpStatus
Public Declare Function GdipCreatePath2 Lib "gdiplus" (Points As POINTF, Types As Any, ByVal Count As Long, BrushMode As FillMode, Path As Long) As GpStatus
Public Declare Function GdipCreatePath2I Lib "gdiplus" (Points As POINTL, Types As Any, ByVal Count As Long, BrushMode As FillMode, Path As Long) As GpStatus
Public Declare Function GdipCreatePath2_ Lib "gdiplus" Alias "GdipCreatePath2" (Points As Any, Types As Any, ByVal Count As Long, BrushMode As FillMode, Path As Long) As GpStatus
Public Declare Function GdipCreatePath2I_ Lib "gdiplus" Alias "GdipCreatePath2I" (Points As Any, Types As Any, ByVal Count As Long, BrushMode As FillMode, Path As Long) As GpStatus
Public Declare Function GdipClonePath Lib "gdiplus" (ByVal Path As Long, clonePath As Long) As GpStatus
Public Declare Function GdipDeletePath Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipResetPath Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPointCount Lib "gdiplus" (ByVal Path As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathTypes Lib "gdiplus" (ByVal Path As Long, Types As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPoints Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPointsI Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPoints_ Lib "gdiplus" Alias "GdipGetPathPoints" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPointsI_ Lib "gdiplus" Alias "GdipGetPathPointsI" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathFillMode Lib "gdiplus" (ByVal Path As Long, ByVal BrushMode As FillMode) As GpStatus
Public Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal Path As Long, ByVal BrushMode As FillMode) As GpStatus
Public Declare Function GdipGetPathData Lib "gdiplus" (ByVal Path As Long, pData As PathData) As GpStatus
Public Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClosePathFigures Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipSetPathMarker Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClearPathMarkers Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipReversePath Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPathLastPoint Lib "gdiplus" (ByVal Path As Long, lastPoint As POINTF) As GpStatus
Public Declare Function GdipAddPathLine Lib "gdiplus" (ByVal Path As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GpStatus
Public Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathLine2_ Lib "gdiplus" Alias "GdipAddPathLine2" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathArc Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezier Lib "gdiplus" (ByVal Path As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As GpStatus
Public Declare Function GdipAddPathBeziers Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathCurve3 Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurve Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathBeziers_ Lib "gdiplus" Alias "GdipAddPathBeziers" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve_ Lib "gdiplus" Alias "GdipAddPathCurve" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2_ Lib "gdiplus" Alias "GdipAddPathCurve2" (ByVal Path As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathCurve3_ Lib "gdiplus" Alias "GdipAddPathCurve3" (ByVal Path As Long, Points As Any, ByVal Count As Long, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurve_ Lib "gdiplus" Alias "GdipAddPathClosedCurve" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2_ Lib "gdiplus" Alias "GdipAddPathClosedCurve2" (ByVal Path As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathRectangles Lib "gdiplus" (ByVal Path As Long, RECT As RECTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathPie Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal Path As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPolygon_ Lib "gdiplus" Alias "GdipAddPathPolygon" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPath Lib "gdiplus" (ByVal Path As Long, ByVal addingPath As Long, ByVal bConnect As Long) As GpStatus
Public Declare Function GdipAddPathString Lib "gdiplus" (ByVal Path As Long, ByVal Str As Long, ByVal Length As Long, ByVal family As Long, ByVal Style As FontStyle, ByVal emSize As Single, layoutRect As RECTF, ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathStringI Lib "gdiplus" (ByVal Path As Long, ByVal Str As Long, ByVal Length As Long, ByVal family As Long, ByVal Style As FontStyle, ByVal emSize As Single, layoutRect As RECTL, ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal Path As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GpStatus
Public Declare Function GdipAddPathLine2I Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathLine2I_ Lib "gdiplus" Alias "GdipAddPathLine2I" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathArcI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezierI Lib "gdiplus" (ByVal Path As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As GpStatus
Public Declare Function GdipAddPathBeziersI Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurveI Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2I Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long, ByVal tension As Long) As GpStatus
Public Declare Function GdipAddPathCurve3I Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurveI Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2I Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathBeziersI_ Lib "gdiplus" Alias "GdipAddPathBeziersI" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurveI_ Lib "gdiplus" Alias "GdipAddPathCurveI" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2I_ Lib "gdiplus" Alias "GdipAddPathCurve2I" (ByVal Path As Long, Points As Any, ByVal Count As Long, ByVal tension As Long) As GpStatus
Public Declare Function GdipAddPathCurve3I_ Lib "gdiplus" Alias "GdipAddPathCurve3I" (ByVal Path As Long, Points As Any, ByVal Count As Long, ByVal offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurveI_ Lib "gdiplus" Alias "GdipAddPathClosedCurveI" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2I_ Lib "gdiplus" Alias "GdipAddPathClosedCurve2I" (ByVal Path As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathRectanglesI Lib "gdiplus" (ByVal Path As Long, Rects As RECTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathEllipseI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathPieI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygonI Lib "gdiplus" (ByVal Path As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPolygonI_ Lib "gdiplus" Alias "GdipAddPathPolygonI" (ByVal Path As Long, Points As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipFlattenPath Lib "gdiplus" (ByVal Path As Long, Optional ByVal Matrix As Long = 0, Optional ByVal Flatness As Single = 0.25) As GpStatus
Public Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal Path As Long, ByVal Matrix As Long, ByVal Flatness As Single) As GpStatus
Public Declare Function GdipWidenPath Lib "gdiplus" (ByVal NativePath As Long, ByVal Pen As Long, ByVal Matrix As Long, ByVal Flatness As Single) As GpStatus
Public Declare Function GdipWarpPath Lib "gdiplus" (ByVal Path As Long, ByVal Matrix As Long, Points As POINTF, ByVal Count As Long, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal WarpMd As WarpMode, ByVal Flatness As Single) As GpStatus
Public Declare Function GdipWarpPath_ Lib "gdiplus" Alias "GdipWarpPath" (ByVal Path As Long, ByVal Matrix As Long, Points As Any, ByVal Count As Long, ByVal SrcX As Single, ByVal SrcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal WarpMd As WarpMode, ByVal Flatness As Single) As GpStatus
Public Declare Function GdipTransformPath Lib "gdiplus" (ByVal Path As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal Path As Long, Bounds As RECTF, ByVal Matrix As Long, ByVal Pen As Long) As GpStatus
Public Declare Function GdipGetPathWorldBoundsI Lib "gdiplus" (ByVal Path As Long, Bounds As RECTL, ByVal Matrix As Long, ByVal Pen As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Pen As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPointI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Pen As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipCreatePathIter Lib "gdiplus" (iterator As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipDeletePathIter Lib "gdiplus" (ByVal iterator As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpath Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, startIndex As Long, endIndex As Long, isClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpathPath Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, ByVal Path As Long, isClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextPathType Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, pathType As Any, startIndex As Long, endIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarker Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, startIndex As Long, endIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarkerPath Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipPathIterGetCount Lib "gdiplus" (ByVal iterator As Long, Count As Long) As GpStatus
Public Declare Function GdipPathIterGetSubpathCount Lib "gdiplus" (ByVal iterator As Long, Count As Long) As GpStatus
Public Declare Function GdipPathIterIsValid Lib "gdiplus" (ByVal iterator As Long, valid As Long) As GpStatus
Public Declare Function GdipPathIterHasCurve Lib "gdiplus" (ByVal iterator As Long, hasCurve As Long) As GpStatus
Public Declare Function GdipPathIterRewind Lib "gdiplus" (ByVal iterator As Long) As GpStatus
Public Declare Function GdipPathIterEnumerate Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, Points As POINTF, Types As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipPathIterCopyData Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, Points As POINTF, Types As Any, ByVal startIndex As Long, ByVal endIndex As Long) As GpStatus
Public Declare Function GdipPathIterEnumerate_ Lib "gdiplus" Alias "GdipPathIterEnumerate" (ByVal iterator As Long, resultCount As Long, Points As Any, Types As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipPathIterCopyData_ Lib "gdiplus" Alias "GdipPathIterCopyData" (ByVal iterator As Long, resultCount As Long, Points As Any, Types As Any, ByVal startIndex As Long, ByVal endIndex As Long) As GpStatus
Public Declare Function GdipCreateMatrix Lib "gdiplus" (Matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix2 Lib "gdiplus" (ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dx As Single, ByVal dy As Single, Matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3 Lib "gdiplus" (RECT As RECTF, dstplg As POINTF, Matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3I Lib "gdiplus" (RECT As RECTL, dstplg As POINTL, Matrix As Long) As GpStatus
Public Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal Matrix As Long, cloneMatrix As Long) As GpStatus
Public Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal Matrix As Long) As GpStatus
Public Declare Function GdipSetMatrixElements Lib "gdiplus" (ByVal Matrix As Long, ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dx As Single, ByVal dy As Single) As GpStatus
Public Declare Function GdipMultiplyMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal matrix2 As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipShearMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal Matrix As Long) As GpStatus
Public Declare Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal Matrix As Long, pts As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPointsI Lib "gdiplus" (ByVal Matrix As Long, pts As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPoints Lib "gdiplus" (ByVal Matrix As Long, pts As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPointsI Lib "gdiplus" (ByVal Matrix As Long, pts As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPoints_ Lib "gdiplus" Alias "GdipTransformMatrixPoints" (ByVal Matrix As Long, pts As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPointsI_ Lib "gdiplus" Alias "GdipTransformMatrixPointsI" (ByVal Matrix As Long, pts As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPoints_ Lib "gdiplus" Alias "GdipVectorTransformMatrixPoints" (ByVal Matrix As Long, pts As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPointsI_ Lib "gdiplus" Alias "GdipVectorTransformMatrixPointsI" (ByVal Matrix As Long, pts As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetMatrixElements Lib "gdiplus" (ByVal Matrix As Long, matrixOut As Single) As GpStatus
Public Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal Matrix As Long, Result As Long) As GpStatus
Public Declare Function GdipIsMatrixIdentity Lib "gdiplus" (ByVal Matrix As Long, Result As Long) As GpStatus
Public Declare Function GdipIsMatrixEqual Lib "gdiplus" (ByVal Matrix As Long, ByVal matrix2 As Long, Result As Long) As GpStatus
Public Declare Function GdipCreateRegion Lib "gdiplus" (region As Long) As GpStatus
Public Declare Function GdipCreateRegionRect Lib "gdiplus" (RECT As RECTF, region As Long) As GpStatus
Public Declare Function GdipCreateRegionRectI Lib "gdiplus" (RECT As RECTL, region As Long) As GpStatus
Public Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal Path As Long, region As Long) As GpStatus
Public Declare Function GdipCreateRegionRgnData Lib "gdiplus" (regionData As Any, ByVal Size As Long, region As Long) As GpStatus
Public Declare Function GdipCreateRegionHrgn Lib "gdiplus" (ByVal hRgn As Long, region As Long) As GpStatus
Public Declare Function GdipCloneRegion Lib "gdiplus" (ByVal region As Long, cloneRegion As Long) As GpStatus
Public Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipSetInfinite Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipSetEmpty Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipCombineRegionRect Lib "gdiplus" (ByVal region As Long, RECT As RECTF, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal region As Long, RECT As RECTL, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal region As Long, ByVal Path As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal region As Long, ByVal region2 As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipTranslateRegion Lib "gdiplus" (ByVal region As Long, ByVal dx As Single, ByVal dy As Single) As GpStatus
Public Declare Function GdipTranslateRegionI Lib "gdiplus" (ByVal region As Long, ByVal dx As Long, ByVal dy As Long) As GpStatus
Public Declare Function GdipTransformRegion Lib "gdiplus" (ByVal region As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal region As Long, ByVal Graphics As Long, RECT As RECTF) As GpStatus
Public Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal region As Long, ByVal Graphics As Long, RECT As RECTL) As GpStatus
Public Declare Function GdipGetRegionHRgn Lib "gdiplus" (ByVal region As Long, ByVal Graphics As Long, hRgn As Long) As GpStatus
Public Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal region As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsInfiniteRegion Lib "gdiplus" (ByVal region As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsEqualRegion Lib "gdiplus" (ByVal region As Long, ByVal region2 As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipGetRegionDataSize Lib "gdiplus" (ByVal region As Long, bufferSize As Long) As GpStatus
Public Declare Function GdipGetRegionData Lib "gdiplus" (ByVal region As Long, Buffer As Any, ByVal bufferSize As Long, sizeFilled As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPoint Lib "gdiplus" (ByVal region As Long, ByVal X As Single, ByVal Y As Single, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPointI Lib "gdiplus" (ByVal region As Long, ByVal X As Long, ByVal Y As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRect Lib "gdiplus" (ByVal region As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRectI Lib "gdiplus" (ByVal region As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipGetRegionScansCount Lib "gdiplus" (ByVal region As Long, Ucount As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScans Lib "gdiplus" (ByVal region As Long, Rects As RECTF, Count As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScansI Lib "gdiplus" (ByVal region As Long, Rects As RECTL, Count As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipCreateImageAttributes Lib "gdiplus" (imageattr As Long) As GpStatus
Public Declare Function GdipCloneImageAttributes Lib "gdiplus" (ByVal imageattr As Long, cloneImageattr As Long) As GpStatus
Public Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As GpStatus
Public Declare Function GdipSetImageAttributesToIdentity Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipResetImageAttributes Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, colourMatrix As Any, grayMatrix As Any, ByVal Flags As ColorMatrixFlags) As GpStatus
Public Declare Function GdipSetImageAttributesThreshold Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal Threshold As Single) As GpStatus
Public Declare Function GdipSetImageAttributesGamma Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal Gamma As Single) As GpStatus
Public Declare Function GdipSetImageAttributesNoOp Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long) As GpStatus
Public Declare Function GdipSetImageAttributesColorKeys Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal colorLow As Long, ByVal colorHigh As Long) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannel Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjstType As ColorAdjustType, ByVal enableFlag As Long, ByVal channelFlags As ColorChannelFlags) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannelColorProfile Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal colorProfileFilename As Long) As GpStatus
Public Declare Function GdipSetImageAttributesRemapTable Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal mapSize As Long, Map As Any) As GpStatus
Public Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal imageattr As Long, ByVal Wrap As WrapMode, ByVal ARGB As Long, ByVal bClamp As Long) As GpStatus
Public Declare Function GdipSetImageAttributesICMMode Lib "gdiplus" (ByVal imageattr As Long, ByVal bOn As Long) As GpStatus
Public Declare Function GdipGetImageAttributesAdjustedPalette Lib "gdiplus" (ByVal imageattr As Long, colorPal As ColorPalette, ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As Long, ByVal fontCollection As Long, fontFamily As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As GpStatus
Public Declare Function GdipCloneFontFamily Lib "gdiplus" (ByVal fontFamily As Long, clonedFontFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySerif Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilyMonospace Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal family As Long, ByVal name As Long, ByVal language As Integer) As GpStatus
Public Declare Function GdipIsStyleAvailable Lib "gdiplus" (ByVal family As Long, ByVal Style As Long, IsStyleAvailable As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerable Lib "gdiplus" (ByVal fontCollection As Long, ByVal Graphics As Long, numFound As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerate Lib "gdiplus" (ByVal fontCollection As Long, ByVal numSought As Long, gpFamilies As Long, ByVal numFound As Long, ByVal Graphics As Long) As GpStatus
Public Declare Function GdipGetEmHeight Lib "gdiplus" (ByVal family As Long, ByVal Style As FontStyle, EmHeight As Integer) As GpStatus
Public Declare Function GdipGetCellAscent Lib "gdiplus" (ByVal family As Long, ByVal Style As FontStyle, CellAscent As Integer) As GpStatus
Public Declare Function GdipGetCellDescent Lib "gdiplus" (ByVal family As Long, ByVal Style As FontStyle, CellDescent As Integer) As GpStatus
Public Declare Function GdipGetLineSpacing Lib "gdiplus" (ByVal family As Long, ByVal Style As FontStyle, LineSpacing As Integer) As GpStatus
Public Declare Function GdipCreateFontFromDC Lib "gdiplus" (ByVal hdc As Long, createdfont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontA Lib "gdiplus" (ByVal hdc As Long, logfont As LOGFONTA, createdfont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontW Lib "gdiplus" (ByVal hdc As Long, logfont As LOGFONTW, createdfont As Long) As GpStatus
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal Style As FontStyle, ByVal Unit As GpUnit, createdfont As Long) As GpStatus
Public Declare Function GdipCloneFont Lib "gdiplus" (ByVal curFont As Long, cloneFont As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As GpStatus
Public Declare Function GdipGetFamily Lib "gdiplus" (ByVal curFont As Long, family As Long) As GpStatus
Public Declare Function GdipGetFontStyle Lib "gdiplus" (ByVal curFont As Long, Style As FontStyle) As GpStatus
Public Declare Function GdipGetFontSize Lib "gdiplus" (ByVal curFont As Long, Size As Single) As GpStatus
Public Declare Function GdipGetFontUnit Lib "gdiplus" (ByVal curFont As Long, Unit As GpUnit) As GpStatus
Public Declare Function GdipGetFontHeight Lib "gdiplus" (ByVal curFont As Long, ByVal Graphics As Long, Height As Single) As GpStatus
Public Declare Function GdipGetFontHeightGivenDPI Lib "gdiplus" (ByVal curFont As Long, ByVal DPI As Single, Height As Single) As GpStatus
Public Declare Function GdipGetLogFontA Lib "gdiplus" (ByVal curFont As Long, ByVal Graphics As Long, logfont As LOGFONTA) As GpStatus
Public Declare Function GdipGetLogFontW Lib "gdiplus" (ByVal curFont As Long, ByVal Graphics As Long, logfont As LOGFONTW) As GpStatus
Public Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal fontCollection As Long, numFound As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal fontCollection As Long, ByVal numSought As Long, gpFamilies As Long, numFound As Long) As GpStatus
Public Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal fontCollection As Long, ByVal FileName As Long) As GpStatus
Public Declare Function GdipPrivateAddMemoryFont Lib "gdiplus" (ByVal fontCollection As Long, ByVal memory As Long, ByVal Length As Long) As GpStatus
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal Graphics As Long, ByVal Str As Long, ByVal Length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal Brush As Long) As GpStatus
Public Declare Function GdipMeasureString Lib "gdiplus" (ByVal Graphics As Long, ByVal Str As Long, ByVal Length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, boundingBox As RECTF, codepointsFitted As Long, linesFilled As Long) As GpStatus
Public Declare Function GdipMeasureCharacterRanges Lib "gdiplus" (ByVal Graphics As Long, ByVal Str As Long, ByVal Length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal regionCount As Long, regions As Long) As GpStatus
Public Declare Function GdipDrawDriverString Lib "gdiplus" (ByVal Graphics As Long, ByVal Str As Long, ByVal Length As Long, ByVal thefont As Long, ByVal Brush As Long, positions As POINTF, ByVal Flags As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString Lib "gdiplus" (ByVal Graphics As Long, ByVal Str As Long, ByVal Length As Long, ByVal thefont As Long, positions As POINTF, ByVal Flags As Long, ByVal Matrix As Long, boundingBox As RECTF) As GpStatus
Public Declare Function GdipDrawDriverString_ Lib "gdiplus" Alias "GdipDrawDriverString" (ByVal Graphics As Long, ByVal Str As Long, ByVal Length As Long, ByVal thefont As Long, ByVal Brush As Long, positions As Any, ByVal Flags As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString_ Lib "gdiplus" Alias "GdipMeasureDriverString" (ByVal Graphics As Long, ByVal Str As Long, ByVal Length As Long, ByVal thefont As Long, positions As Any, ByVal Flags As Long, ByVal Matrix As Long, boundingBox As RECTF) As GpStatus
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericDefault Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericTypographic Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipCloneStringFormat Lib "gdiplus" (ByVal StringFormat As Long, newFormat As Long) As GpStatus
Public Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal StringFormat As Long, ByVal Flags As Long) As GpStatus
Public Declare Function GdipGetStringFormatFlags Lib "gdiplus" (ByVal StringFormat As Long, Flags As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, ByVal Trimming As StringTrimming) As GpStatus
Public Declare Function GdipGetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, Trimming As Long) As GpStatus
Public Declare Function GdipSetStringFormatHotkeyPrefix Lib "gdiplus" (ByVal StringFormat As Long, ByVal hkPrefix As GpHotkeyPrefix) As GpStatus
Public Declare Function GdipGetStringFormatHotkeyPrefix Lib "gdiplus" (ByVal StringFormat As Long, hkPrefix As GpHotkeyPrefix) As GpStatus
Public Declare Function GdipSetStringFormatTabStops Lib "gdiplus" (ByVal StringFormat As Long, ByVal firstTabOffset As Single, ByVal Count As Long, tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStops Lib "gdiplus" (ByVal StringFormat As Long, ByVal Count As Long, firstTabOffset As Single, tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStopCount Lib "gdiplus" (ByVal StringFormat As Long, Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatDigitSubstitution Lib "gdiplus" (ByVal StringFormat As Long, ByVal language As Integer, ByVal substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatDigitSubstitution Lib "gdiplus" (ByVal StringFormat As Long, language As Integer, substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatMeasurableCharacterRangeCount Lib "gdiplus" (ByVal StringFormat As Long, Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatMeasurableCharacterRanges Lib "gdiplus" (ByVal StringFormat As Long, ByVal rangeCount As Long, ranges As CharacterRange) As GpStatus
Public Declare Function GdipCreateEffect Lib "gdiplus" (ByVal Guid41 As Long, ByVal Guid42 As Long, ByVal Guid43 As Long, ByVal Guid44 As Long, Effect As Long) As GpStatus
Public Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal Effect As Long) As GpStatus
Public Declare Function GdipGetEffectParameterSize Lib "gdiplus" (ByVal Effect As Long, Size As Long) As GpStatus
Public Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal Effect As Long, Params As Any, ByVal Size As Long) As GpStatus
Public Declare Function GdipGetEffectParameters Lib "gdiplus" (ByVal Effect As Long, Size As Long, Params As Any) As GpStatus
Public Declare Function GdipImageSetAbort Lib "gdiplus" (ByVal Image As Long, IAbort As GdiplusAbort) As GpStatus
Public Declare Function GdipGraphicsSetAbort Lib "gdiplus" (ByVal Graphics As Long, IAbort As GdiplusAbort) As GpStatus
Public Declare Function GdipBitmapConvertFormat Lib "gdiplus" (ByVal InputBitmap As Long, ByVal Format As GpPixelFormat, ByVal DitherType As DitherType, ByVal PaletteType As PaletteType, Palette As ColorPalette, ByVal alphaThresholdPercent As Single) As GpStatus
Public Declare Function GdipInitializePalette Lib "gdiplus" (Palette As ColorPalette, ByVal PaletteType As PaletteType, ByVal optimalColors As Long, ByVal useTransparentColor As Long, Optional ByVal Bitmap As Long) As GpStatus
Public Declare Function GdipBitmapApplyEffect Lib "gdiplus" (ByVal Bitmap As Long, ByVal Effect As Long, roi As RECTL, ByVal useAuxData As Long, auxData As Any, auxDataSize As Long) As GpStatus
Public Declare Function GdipBitmapCreateApplyEffect Lib "gdiplus" (inputBitmaps As Any, ByVal numInputs As Long, ByVal Effect As Long, roi As RECTL, outputRect As RECTL, outputBitmap As Long, ByVal useAuxData As Long, auxData As Any, auxDataSize As Long) As GpStatus
Public Declare Function GdipBitmapGetHistogram Lib "gdiplus" (ByVal Bitmap As Long, ByVal Format As HistogramFormat, ByVal NumberOfEntries As Long, channel0 As Any, channel1 As Any, channel2 As Any, channel3 As Any) As GpStatus
Public Declare Function GdipBitmapGetHistogramSize Lib "gdiplus" (ByVal Format As HistogramFormat, NumberOfEntries As Long) As GpStatus
Public Declare Function GdipFindFirstImageItem Lib "gdiplus" (ByVal Image As Long, Item As ImageItemData) As GpStatus
Public Declare Function GdipFindNextImageItem Lib "gdiplus" (ByVal Image As Long, Item As ImageItemData) As GpStatus
Public Declare Function GdipGetImageItemData Lib "gdiplus" (ByVal Image As Long, Item As ImageItemData) As GpStatus
Public Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, Source As RECTF, ByVal xForm As Long, ByVal Effect As Long, ByVal imageAttributes As Long, ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipCreateFromHDC2 Lib "gdiplus" (ByVal hdc As Long, ByVal hDevice As Long, Graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWNDICM Lib "gdiplus" (ByVal Hwnd As Long, Graphics As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPoint Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destPoint As POINTF, lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPointI Lib "gdiplus" (Graphics As Long, ByVal metafile As Long, destPoint As POINTL, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRect Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destRect As RECTF, lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destRect As RECTL, lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPoints Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destPoint As POINTF, ByVal Count As Long, lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destPoint As POINTL, ByVal Count As Long, lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoint Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destPoint As POINTF, srcRect As RECTF, ByVal srcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointI Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destPoint As POINTL, srcRect As RECTL, ByVal srcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRect Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destRect As RECTF, srcRect As RECTF, ByVal srcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destRect As RECTL, srcRect As RECTL, ByVal srcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoints Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destPoints As POINTF, ByVal Count As Long, srcRect As RECTF, ByVal srcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal metafile As Long, destPoints As POINTL, ByVal Count As Long, srcRect As RECTL, ByVal srcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoints_ Lib "gdiplus" Alias "GdipEnumerateMetafileSrcRectDestPoints" (ByVal Graphics As Long, ByVal metafile As Long, destPoints As Any, ByVal Count As Long, srcRect As RECTF, ByVal srcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointsI_ Lib "gdiplus" Alias "GdipEnumerateMetafileSrcRectDestPointsI" (ByVal Graphics As Long, ByVal metafile As Long, destPoints As Any, ByVal Count As Long, srcRect As RECTL, ByVal srcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipPlayMetafileRecord Lib "gdiplus" (ByVal metafile As Long, ByVal recordType As EmfPlusRecordType, ByVal Flags As Long, ByVal dataSize As Long, byteData As Any) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromWmf Lib "gdiplus" (ByVal hWmf As Long, WmfPlaceableFileHdr As WmfPlaceableFileHeader, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromEmf Lib "gdiplus" (ByVal hEmf As Long, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromFile Lib "gdiplus" (ByVal FileName As Long, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromStream Lib "gdiplus" (ByVal stream As Any, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromMetafile Lib "gdiplus" (ByVal metafile As Long, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetHemfFromMetafile Lib "gdiplus" (ByVal metafile As Long, hEmf As Long) As GpStatus
Public Declare Function GdipCreateStreamOnFile Lib "gdiplus" (ByVal FileName As Long, ByVal access As Long, stream As Any) As GpStatus
Public Declare Function GdipCreateMetafileFromWmf Lib "gdiplus" (ByVal hWmf As Long, ByVal bDeleteWmf As Long, WmfPlaceableFileHdr As WmfPlaceableFileHeader, ByVal metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromEmf Lib "gdiplus" (ByVal hEmf As Long, ByVal bDeleteEmf As Long, metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromFile Lib "gdiplus" (ByVal File As Long, metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromWmfFile Lib "gdiplus" (ByVal File As Long, WmfPlaceableFileHdr As WmfPlaceableFileHeader, metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromStream Lib "gdiplus" (ByVal stream As Any, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafile Lib "gdiplus" (ByVal referenceHdc As Long, etype As emfType, frameRect As RECTF, ByVal frameUnit As MetafileFrameUnit, ByVal Description As Long, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileI Lib "gdiplus" (ByVal referenceHdc As Long, etype As emfType, frameRect As RECTL, ByVal frameUnit As MetafileFrameUnit, ByVal Description As Long, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileName Lib "gdiplus" (ByVal FileName As Long, ByVal referenceHdc As Long, etype As emfType, frameRect As RECTF, ByVal frameUnit As MetafileFrameUnit, ByVal Description As Long, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileNameI Lib "gdiplus" (ByVal FileName As Long, ByVal referenceHdc As Long, etype As emfType, frameRect As RECTL, ByVal frameUnit As MetafileFrameUnit, ByVal Description As Long, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileStream Lib "gdiplus" (ByVal stream As Any, ByVal referenceHdc As Long, etype As emfType, frameRect As RECTF, ByVal frameUnit As MetafileFrameUnit, ByVal Description As Long, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileStreamI Lib "gdiplus" (ByVal stream As Any, ByVal referenceHdc As Long, etype As emfType, frameRect As RECTL, ByVal frameUnit As MetafileFrameUnit, ByVal Description As Long, metafile As Long) As GpStatus
Public Declare Function GdipSetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal metafile As Long, ByVal metafileRasterizationLimitDpi As Long) As GpStatus
Public Declare Function GdipGetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal metafile As Long, metafileRasterizationLimitDpi As Long) As GpStatus
Public Declare Function GdipGetImageDecodersSize Lib "gdiplus" (numDecoders As Long, Size As Long) As GpStatus
Public Declare Function GdipSetImageAttributesCachedBackground Lib "gdiplus" (ByVal imageattr As Long, ByVal enableFlag As Long) As GpStatus
Public Declare Function GdipTestControl Lib "gdiplus" (ByVal Control As GpTestControlEnum, param As Any) As GpStatus
Public Declare Function GdipConvertToEmfPlus Lib "gdiplus" (ByVal refGraphics As Long, conversionFailureFlag As Long, ByVal emfType As emfType, ByVal Description As Long, ByVal out_metafile As Long) As GpStatus
Public Declare Function GdipConvertToEmfPlusToFile Lib "gdiplus" (ByVal refGraphics As Long, ByVal metafile As Long, conversionFailureFlag As Long, ByVal FileName As Long, ByVal emfType As emfType, ByVal Description As Long, out_metafile As Long) As GpStatus
Public Declare Function GdipConvertToEmfPlusToStream Lib "gdiplus" (ByVal refGraphics As Long, ByVal metafile As Long, conversionFailureFlag As Long, stream As Any, ByVal emfType As emfType, ByVal Description As Long, out_metafile As Long) As GpStatus
Public Declare Function GdipFlush Lib "gdiplus" (ByVal Graphics As Long, ByVal intention As FlushIntention) As GpStatus
Public Declare Function GdipAlloc Lib "gdiplus" (ByVal Size As Long) As Long
Public Declare Sub GdipFree Lib "gdiplus" (ByVal ptr As Long)
Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, Inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus