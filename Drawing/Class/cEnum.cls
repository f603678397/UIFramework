VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnum"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum GpPixelFormat
    PixelFormat1bppIndexed = &H30101
    PixelFormat4bppIndexed = &H30402
    PixelFormat8bppIndexed = &H30803
    PixelFormat16bppGreyScale = &H101004
    PixelFormat16bppRGB555 = &H21005
    PixelFormat16bppRGB565 = &H21006
    PixelFormat16bppARGB1555 = &H61007
    PixelFormat24bppRGB = &H21808
    PixelFormat32bppRGB = &H22009
    PixelFormat32bppARGB = &H26200A
    PixelFormat32bppPARGB = &HE200B
    PixelFormat48bppRGB = &H10300C
    PixelFormat64bppARGB = &H34400D
    PixelFormat64bppPARGB = &H1C400E
End Enum
'==================================
Public Enum GpUnit
    UnitWorld = 0
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum
'==================================
Public Enum PathPointType
    PathPointTypeStart = 0
    PathPointTypeLine = 1
    PathPointTypeBezier = 3
    PathPointTypePathTypeMask = &H7
    PathPointTypePathDashMode = &H10
    PathPointTypePathMarker = &H20
    PathPointTypeCloseSubpath = &H80
    PathPointTypeBezier3 = 3
End Enum
'==================================
Public Enum GenericFontFamily
    GenericFontFamilySerif = 0
    GenericFontFamilySansSerif
    GenericFontFamilyMonospace
End Enum

Public Enum FontStyle
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Public Enum StringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter = 1
    StringAlignmentFar = 2
End Enum
'==================================
Public Enum FillMode
    FillModeAlternate = 0
    FillModeWinding
End Enum

Public Enum WrapMode
    WrapModeTile = 0
    WrapModeTileFlipX
    WrapModeTileFlipy
    WrapModeTileFlipXY
    WrapModeClamp
End Enum

Public Enum LinearGradientMode
    LinearGradientModeHorizontal = 0
    LinearGradientModeVertical
    LinearGradientModeForwardDiagonal
    LinearGradientModeBackwardDiagonal
End Enum
'==================================
Public Enum QualityMode
    QualityModeInvalid = -1
    QualityModeDefault = 0
    QualityModeLow = 1
    QualityModeHigh = 2
End Enum

Public Enum CompositingMode
    CompositingModeSourceOver = 0
    CompositingModeSourceCopy
End Enum

Public Enum CompositingQuality
    CompositingQualityInvalid = QualityModeInvalid
    CompositingQualityDefault = QualityModeDefault
    CompositingQualityHighSpeed = QualityModeLow
    CompositingQualityHighQuality = QualityModeHigh
    CompositingQualityGammaCorrected
    CompositingQualityAssumeLinear
End Enum

Public Enum SmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone
    SmoothingModeAntiAlias
    SmoothingModeAntiAlias8x4 = SmoothingModeAntiAlias
    SmoothingModeAntiAlias8x8
End Enum

Public Enum InterpolationMode
    InterpolationModeInvalid = QualityModeInvalid
    InterpolationModeDefault = QualityModeDefault
    InterpolationModeLowQuality = QualityModeLow
    InterpolationModeHighQuality = QualityModeHigh
    InterpolationModeBilinear
    InterpolationModeBicubic
    InterpolationModeNearestNeighbor
    InterpolationModeHighQualityBilinear
    InterpolationModeHighQualityBicubic
End Enum

Public Enum PixelOffsetMode
    PixelOffsetModeInvalid = QualityModeInvalid
    PixelOffsetModeDefault = QualityModeDefault
    PixelOffsetModeHighSpeed = QualityModeLow
    PixelOffsetModeHighQuality = QualityModeHigh
    PixelOffsetModeNone
    PixelOffsetModeHalf
End Enum

Public Enum TextRenderingHint
    TextRenderingHintSystemDefault = 0
    TextRenderingHintSingleBitPerPixelGridFit
    TextRenderingHintSingleBitPerPixel
    TextRenderingHintAntiAliasGridFit
    TextRenderingHintAntiAlias
    TextRenderingHintClearTypeGridFit
End Enum
'==================================
Public Enum MatrixOrder
    MatrixOrderPrepend = 0
    MatrixOrderAppend = 1
End Enum

Public Enum ColorAdjustType
    ColorAdjustTypeDefault = 0
    ColorAdjustTypeBitmap
    ColorAdjustTypeBrush
    ColorAdjustTypePen
    ColorAdjustTypeText
    ColorAdjustTypeCount
    ColorAdjustTypeAny
End Enum

Public Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0
    ColorMatrixFlagsSkipGrays = 1
    ColorMatrixFlagsAltGray = 2
End Enum

Public Enum WarpMode
    WarpModePerspective = 0
    WarpModeBilinear
End Enum

Public Enum CombineMode
    CombineModeReplace = 0
    CombineModeIntersect
    CombineModeUnion
    CombineModeXor
    CombineModeExclude
    CombineModeComplement
End Enum

Public Enum ImageLockMode
    ImageLockModeRead = 1
    ImageLockModeWrite = 2
    ImageLockModeUserInputBuf = 4
End Enum
'==================================
Public Enum GdipEffectType
    Blur = 0
    BrightnessContrast = 1
    ColorBalance = 2
    ColorCurve = 3
    ColorLookupTable = 4
    ColorMatrix = 5
    HueSaturationLightness = 6
    levels = 7
    RedEyeCorrection = 8
    Sharpen = 9
    Tint = 10
End Enum

Public Enum HistogramFormat
    HistogramFormatARGB
    HistogramFormatPARGB
    HistogramFormatRGB
    HistogramFormatGray
    HistogramFormatB
    HistogramFormatG
    HistogramFormatR
    HistogramFormatA
End Enum

Public Enum CurveAdjustments
    AdjustExposure
    AdjustDensity
    AdjustContrast
    AdjustHighlight
    AdjustShadow
    AdjustMidtone
    AdjustWhiteSaturation
    AdjustBlackSaturation
End Enum

Public Enum CurveChannel
    CurveChannelAll
    CurveChannelRed
    CurveChannelGreen
    CurveChannelBlue
End Enum

Public Enum PaletteType
    PaletteTypeCustom = 0
    PaletteTypeOptimal = 1
    PaletteTypeFixedBW = 2
    PaletteTypeFixedHalftone8 = 3
    PaletteTypeFixedHalftone27 = 4
    PaletteTypeFixedHalftone64 = 5
    PaletteTypeFixedHalftone125 = 6
    PaletteTypeFixedHalftone216 = 7
    PaletteTypeFixedHalftone252 = 8
    PaletteTypeFixedHalftone256 = 9
End Enum

Public Enum DitherType
    DitherTypeNone = 0
    DitherTypeSolid = 1
    DitherTypeOrdered4x4 = 2
    DitherTypeOrdered8x8 = 3
    DitherTypeOrdered16x16 = 4
    DitherTypeSpiral4x4 = 5
    DitherTypeSpiral8x8 = 6
    DitherTypeDualSpiral4x4 = 7
    DitherTypeDualSpiral8x8 = 8
    DitherTypeErrorDiffusion = 9
    DitherTypeMax = 10
End Enum

Public Enum ItemDataPosition
    ItemDataPositionAfterHeader = 0
    ItemDataPositionAfterPalette = 1
    ItemDataPositionAfterBits = 2
End Enum

Public Enum GpImageSaveFormat
    GpSaveBMP = 0
    GpSaveJPEG = 1
    GpSaveGIF = 2
    GpSavePNG = 3
    GpSaveTIFF = 4
End Enum

Public Enum GpImageFormatIdentifiers
    GpImageFormatUndefined = 0
    GpImageFormatMemoryBMP = 1
    GpImageFormatBMP = 2
    GpImageFormatEMF = 3
    GpImageFormatWMF = 4
    GpImageFormatJPEG = 5
    GpImageFormatPNG = 6
    GpImageFormatGIF = 7
    GpImageFormatTIFF = 8
    GpImageFormatEXIF = 9
    GpImageFormatIcon = 10
End Enum

Public Enum Image_Type
    ImageTypeUnknown = 0
    ImageTypeBitmap = 1
    ImageTypeMetafile = 2
End Enum

Public Enum Image_Property_Types
    PropertyTagTypeByte = 1
    PropertyTagTypeASCII = 2
    PropertyTagTypeShort = 3
    PropertyTagTypeLong = 4
    PropertyTagTypeRational = 5
    PropertyTagTypeUndefined = 7
    PropertyTagTypeSLONG = 9
    PropertyTagTypeSRational = 10
End Enum

Public Enum ImageCodecFlags
    ImageCodecFlagsEncoder = &H1
    ImageCodecFlagsDecoder = &H2
    ImageCodecFlagsSupportBitmap = &H4
    ImageCodecFlagsSupportVector = &H8
    ImageCodecFlagsSeekableEncode = &H10
    ImageCodecFlagsBlockingDecode = &H20
    
    ImageCodecFlagsBuiltin = &H10000
    ImageCodecFlagsSystem = &H20000
    ImageCodecFlagsUser = &H40000
End Enum

Public Enum Image_Property_ID_Tags
    PropertyTagExifIFD = &H8769
    PropertyTagGpsIFD = &H8825
    PropertyTagNewSubfileType = &HFE
    PropertyTagSubfileType = &HFF
    PropertyTagImageWidth = &H100
    PropertyTagImageHeight = &H101
    PropertyTagBitsPerSample = &H102
    PropertyTagCompression = &H103
    PropertyTagPhotometricInterp = &H106
    PropertyTagThreshHolding = &H107
    PropertyTagCellWidth = &H108
    PropertyTagCellHeight = &H109
    PropertyTagFillOrder = &H10A
    PropertyTagDocumentName = &H10D
    PropertyTagImageDescription = &H10E
    PropertyTagEquipMake = &H10F
    PropertyTagEquipModel = &H110
    PropertyTagStripOffsets = &H111
    PropertyTagOrientation = &H112
    PropertyTagSamplesPerPixel = &H115
    PropertyTagRowsPerStrip = &H116
    PropertyTagStripBytesCount = &H117
    PropertyTagMinSampleValue = &H118
    PropertyTagMaxSampleValue = &H119
    PropertyTagXResolution = &H11A
    PropertyTagYResolution = &H11B
    PropertyTagPlanarConfig = &H11C
    PropertyTagPageName = &H11D
    PropertyTagXPosition = &H11E
    PropertyTagYPosition = &H11F
    PropertyTagFreeOffset = &H120
    PropertyTagFreeByteCounts = &H121
    PropertyTagGrayResponseUnit = &H122
    PropertyTagGrayResponseCurve = &H123
    PropertyTagT4Option = &H124
    PropertyTagT6Option = &H125
    PropertyTagResolutionUnit = &H128
    PropertyTagPageNumber = &H129
    PropertyTagTransferFuncition = &H12D
    PropertyTagSoftwareUsed = &H131
    PropertyTagDateTime = &H132
    PropertyTagArtist = &H13B
    PropertyTagHostComputer = &H13C
    PropertyTagPredictor = &H13D
    PropertyTagWhitePoint = &H13E
    PropertyTagPrimaryChromaticities = &H13F
    PropertyTagColorMap = &H140
    PropertyTagHalftoneHints = &H141
    PropertyTagTileWidth = &H142
    PropertyTagTileLength = &H143
    PropertyTagTileOffset = &H144
    PropertyTagTileByteCounts = &H145
    PropertyTagInkSet = &H14C
    PropertyTagInkNames = &H14D
    PropertyTagNumberOfInks = &H14E
    PropertyTagDotRange = &H150
    PropertyTagTargetPrinter = &H151
    PropertyTagExtraSamples = &H152
    PropertyTagSampleFormat = &H153
    PropertyTagSMinSampleValue = &H154
    PropertyTagSMaxSampleValue = &H155
    PropertyTagTransferRange = &H156
    PropertyTagJPEGProc = &H200
    PropertyTagJPEGInterFormat = &H201
    PropertyTagJPEGInterLength = &H202
    PropertyTagJPEGRestartInterval = &H203
    PropertyTagJPEGLosslessPredictors = &H205
    PropertyTagJPEGPointTransforms = &H206
    PropertyTagJPEGQTables = &H207
    PropertyTagJPEGDCTables = &H208
    PropertyTagJPEGACTables = &H209
    PropertyTagYCbCrCoefficients = &H211
    PropertyTagYCbCrSubsampling = &H212
    PropertyTagYCbCrPositioning = &H213
    PropertyTagREFBlackWhite = &H214
    PropertyTagICCProfile = &H8773
    PropertyTagGamma = &H301
    PropertyTagICCProfileDescriptor = &H302
    PropertyTagSRGBRenderingIntent = &H303
    PropertyTagImageTitle = &H320
    PropertyTagCopyright = &H8298
    PropertyTagResolutionXUnit = &H5001
    PropertyTagResolutionYUnit = &H5002
    PropertyTagResolutionXLengthUnit = &H5003
    PropertyTagResolutionYLengthUnit = &H5004
    PropertyTagPrintFlags = &H5005
    PropertyTagPrintFlagsVersion = &H5006
    PropertyTagPrintFlagsCrop = &H5007
    PropertyTagPrintFlagsBleedWidth = &H5008
    PropertyTagPrintFlagsBleedWidthScale = &H5009
    PropertyTagHalftoneLPI = &H500A
    PropertyTagHalftoneLPIUnit = &H500B
    PropertyTagHalftoneDegree = &H500C
    PropertyTagHalftoneShape = &H500D
    PropertyTagHalftoneMisc = &H500E
    PropertyTagHalftoneScreen = &H500F
    PropertyTagJPEGQuality = &H5010
    PropertyTagGridSize = &H5011
    PropertyTagThumbnailFormat = &H5012
    PropertyTagThumbnailWidth = &H5013
    PropertyTagThumbnailHeight = &H5014
    PropertyTagThumbnailColorDepth = &H5015
    PropertyTagThumbnailPlanes = &H5016
    PropertyTagThumbnailRawBytes = &H5017
    PropertyTagThumbnailSize = &H5018
    PropertyTagThumbnailCompressedSize = &H5019
    PropertyTagColorTransferFunction = &H501A
    PropertyTagThumbnailData = &H501B
    PropertyTagThumbnailImageWidth = &H5020
    PropertyTagThumbnailImageHeight = &H5021
    PropertyTagThumbnailBitsPerSample = &H5022
    PropertyTagThumbnailCompression = &H5023
    PropertyTagThumbnailPhotometricInterp = &H5024
    PropertyTagThumbnailImageDescription = &H5025
    PropertyTagThumbnailEquipMake = &H5026
    PropertyTagThumbnailEquipModel = &H5027
    PropertyTagThumbnailStripOffsets = &H5028
    PropertyTagThumbnailOrientation = &H5029
    PropertyTagThumbnailSamplesPerPixel = &H502A
    PropertyTagThumbnailRowsPerStrip = &H502B
    PropertyTagThumbnailStripBytesCount = &H502C
    PropertyTagThumbnailResolutionX = &H502D
    PropertyTagThumbnailResolutionY = &H502E
    PropertyTagThumbnailPlanarConfig = &H502F
    PropertyTagThumbnailResolutionUnit = &H5030
    PropertyTagThumbnailTransferFunction = &H5031
    PropertyTagThumbnailSoftwareUsed = &H5032
    PropertyTagThumbnailDateTime = &H5033
    PropertyTagThumbnailArtist = &H5034
    PropertyTagThumbnailWhitePoint = &H5035
    PropertyTagThumbnailPrimaryChromaticities = &H5036
    PropertyTagThumbnailYCbCrCoefficients = &H5037
    PropertyTagThumbnailYCbCrSubsampling = &H5038
    PropertyTagThumbnailYCbCrPositioning = &H5039
    PropertyTagThumbnailRefBlackWhite = &H503A
    PropertyTagThumbnailCopyRight = &H503B
    PropertyTagLuminanceTable = &H5090
    PropertyTagChrominanceTable = &H5091
    PropertyTagFrameDelay = &H5100
    PropertyTagLoopCount = &H5101
    PropertyTagGlobalPalette = &H5102
    PropertyTagIndexBackground = &H5103
    PropertyTagIndexTransparent = &H5104
    PropertyTagPixelUnit = &H5110
    PropertyTagPixelPerUnitX = &H5111
    PropertyTagPixelPerUnitY = &H5112
    PropertyTagPaletteHistogram = &H5113
    PropertyTagExifExposureTime = &H829A
    PropertyTagExifFNumber = &H829D
    PropertyTagExifExposureProg = &H8822
    PropertyTagExifSpectralSense = &H8824
    PropertyTagExifISOSpeed = &H8827
    PropertyTagExifOECF = &H8828
    PropertyTagExifVer = &H9000
    PropertyTagExifDTOrig = &H9003
    PropertyTagExifDTDigitized = &H9004
    PropertyTagExifCompConfig = &H9101
    PropertyTagExifCompBPP = &H9102
    PropertyTagExifShutterSpeed = &H9201
    PropertyTagExifAperture = &H9202
    PropertyTagExifBrightness = &H9203
    PropertyTagExifExposureBias = &H9204
    PropertyTagExifMaxAperture = &H9205
    PropertyTagExifSubjectDist = &H9206
    PropertyTagExifMeteringMode = &H9207
    PropertyTagExifLightSource = &H9208
    PropertyTagExifFlash = &H9209
    PropertyTagExifFocalLength = &H920A
    PropertyTagExifMakerNote = &H927C
    PropertyTagExifUserComment = &H9286
    PropertyTagExifDTSubsec = &H9290
    PropertyTagExifDTOrigSS = &H9291
    PropertyTagExifDTDigSS = &H9292
    PropertyTagExifFPXVer = &HA000
    PropertyTagExifColorSpace = &HA001
    PropertyTagExifPixXDim = &HA002
    PropertyTagExifPixYDim = &HA003
    PropertyTagExifRelatedWav = &HA004
    PropertyTagExifInterop = &HA005
    PropertyTagExifFlashEnergy = &HA20B
    PropertyTagExifSpatialFR = &HA20C
    PropertyTagExifFocalXRes = &HA20E
    PropertyTagExifFocalYRes = &HA20F
    PropertyTagExifFocalResUnit = &HA210
    PropertyTagExifSubjectLoc = &HA214
    PropertyTagExifExposureIndex = &HA215
    PropertyTagExifSensingMethod = &HA217
    PropertyTagExifFileSource = &HA300
    PropertyTagExifSceneType = &HA301
    PropertyTagExifCfaPattern = &HA302
    PropertyTagGpsVer = &H0
    PropertyTagGpsLatitudeRef = &H1
    PropertyTagGpsLatitude = &H2
    PropertyTagGpsLongitudeRef = &H3
    PropertyTagGpsLongitude = &H4
    PropertyTagGpsAltitudeRef = &H5
    PropertyTagGpsAltitude = &H6
    PropertyTagGpsGpsTime = &H7
    PropertyTagGpsGpsSatellites = &H8
    PropertyTagGpsGpsStatus = &H9
    PropertyTagGpsGpsMeasureMode = &HA
    PropertyTagGpsGpsDop = &HB
    PropertyTagGpsSpeedRef = &HC
    PropertyTagGpsSpeed = &HD
    PropertyTagGpsTrackRef = &HE
    PropertyTagGpsTrack = &HF
    PropertyTagGpsImgDirRef = &H10
    PropertyTagGpsImgDir = &H11
    PropertyTagGpsMapDatum = &H12
    PropertyTagGpsDestLatRef = &H13
    PropertyTagGpsDestLat = &H14
    PropertyTagGpsDestLongRef = &H15
    PropertyTagGpsDestLong = &H16
    PropertyTagGpsDestBearRef = &H17
    PropertyTagGpsDestBear = &H18
    PropertyTagGpsDestDistRef = &H19
    PropertyTagGpsDestDist = &H1A
End Enum
'=================================

Public Enum PaletteFlags
    PaletteFlagsHasAlpha = &H1
    PaletteFlagsGrayScale = &H2
    PaletteFlagsHalftone = &H4
End Enum
'=================================

Public Enum RotateFlipType
    RotateNoneFlipNone = 0
    Rotate90FlipNone = 1
    Rotate180FlipNone = 2
    Rotate270FlipNone = 3

    RotateNoneFlipX = 4
    Rotate90FlipX = 5
    Rotate180FlipX = 6
    Rotate270FlipX = 7

    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipY = Rotate90FlipX

    RotateNoneFlipXY = Rotate180FlipNone
    Rotate90FlipXY = Rotate270FlipNone
    Rotate180FlipXY = RotateNoneFlipNone
    Rotate270FlipXY = Rotate90FlipNone
End Enum
'=================================

Public Enum ColorMode
    ColorModeARGB32 = 0
    ColorModeARGB64 = 1
End Enum

Public Enum ColorChannelFlags
    ColorChannelFlagsC = 0
    ColorChannelFlagsM
    ColorChannelFlagsY
    ColorChannelFlagsK
    ColorChannelFlagsLast
End Enum

Public Enum ColorShiftComponents
    AlphaShift = 24
    RedShift = 16
    GreenShift = 8
    BlueShift = 0
End Enum

Public Enum ColorMaskComponents
    AlphaMask = &HFF000000
    RedMask = &HFF0000
    GreenMask = &HFF00
    BlueMask = &HFF
End Enum

'=================================
Public Enum StringFormatFlags
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Public Enum StringTrimming
    StringTrimmingNone = 0
    StringTrimmingCharacter = 1
    StringTrimmingWord = 2
    StringTrimmingEllipsisCharacter = 3
    StringTrimmingEllipsisWord = 4
    StringTrimmingEllipsisPath = 5
End Enum

Public Enum StringDigitSubstitute
    StringDigitSubstituteUser = 0
    StringDigitSubstituteNone = 1
    StringDigitSubstituteNational = 2
    StringDigitSubstituteTraditional = 3
End Enum
'=================================

Public Enum HatchStyle
    HatchStyleHorizontal                   ' 0
    HatchStyleVertical                     ' 1
    HatchStyleForwardDiagonal              ' 2
    HatchStyleBackwardDiagonal             ' 3
    HatchStyleCross                        ' 4
    HatchStyleDiagonalCross                ' 5
    HatchStyle05Percent                    ' 6
    HatchStyle10Percent                    ' 7
    HatchStyle20Percent                    ' 8
    HatchStyle25Percent                    ' 9
    HatchStyle30Percent                    ' 10
    HatchStyle40Percent                    ' 11
    HatchStyle50Percent                    ' 12
    HatchStyle60Percent                    ' 13
    HatchStyle70Percent                    ' 14
    HatchStyle75Percent                    ' 15
    HatchStyle80Percent                    ' 16
    HatchStyle90Percent                    ' 17
    HatchStyleLightDownwardDiagonal        ' 18
    HatchStyleLightUpwardDiagonal          ' 19
    HatchStyleDarkDownwardDiagonal         ' 20
    HatchStyleDarkUpwardDiagonal           ' 21
    HatchStyleWideDownwardDiagonal         ' 22
    HatchStyleWideUpwardDiagonal           ' 23
    HatchStyleLightVertical                ' 24
    HatchStyleLightHorizontal              ' 25
    HatchStyleNarrowVertical               ' 26
    HatchStyleNarrowHorizontal             ' 27
    HatchStyleDarkVertical                 ' 28
    HatchStyleDarkHorizontal               ' 29
    HatchStyleDashedDownwardDiagonal       ' 30
    HatchStyleDashedUpwardDiagonal         ' 31
    HatchStyleDashedHorizontal             ' 32
    HatchStyleDashedVertical               ' 33
    HatchStyleSmallConfetti                ' 34
    HatchStyleLargeConfetti                ' 35
    HatchStyleZigZag                       ' 36
    HatchStyleWave                         ' 37
    HatchStyleDiagonalBrick                ' 38
    HatchStyleHorizontalBrick              ' 39
    HatchStyleWeave                        ' 40
    HatchStylePlaid                        ' 41
    HatchStyleDivot                        ' 42
    HatchStyleDottedGrid                   ' 43
    HatchStyleDottedDiamond                ' 44
    HatchStyleShingle                      ' 45
    HatchStyleTrellis                      ' 46
    HatchStyleSphere                       ' 47
    HatchStyleSmallGrid                    ' 48
    HatchStyleSmallCheckerBoard            ' 49
    HatchStyleLargeCheckerBoard            ' 50
    HatchStyleOutlinedDiamond              ' 51
    HatchStyleSolidDiamond                 ' 52
    HatchStyleTotal
    HatchStyleLargeGrid = HatchStyleCross  ' 4
    HatchStyleMin = HatchStyleHorizontal
    HatchStyleMax = HatchStyleTotal - 1
End Enum

Public Enum PenAlignment
    PenAlignmentCenter = 0
    PenAlignmentInset = 1
End Enum

Public Enum BrushType
    BrushTypeSolidColor = 0
    BrushTypeHatchFill = 1
    BrushTypeTextureFill = 2
    BrushTypePathGradient = 3
    BrushTypeLinearGradient = 4
End Enum

Public Enum DashStyle
    DashStyleSolid
    DashStyleDash
    DashStyleDot
    DashStyleDashDot
    DashStyleDashDotDot
    DashStyleCustom
End Enum

Public Enum DashCap
    DashCapFlat = 0
    DashCapRound = 2
    DashCapTriangle = 3
End Enum

Public Enum LineCap
    LineCapFlat = 0
    LineCapSquare = 1
    LineCapRound = 2
    LineCapTriangle = 3
    LineCapNoAnchor = &H10
    LineCapSquareAnchor = &H11
    LineCapRoundAnchor = &H12
    LineCapDiamondAnchor = &H13
    LineCapArrowAnchor = &H14
    LineCapCustom = &HFF
    LineCapAnchorMask = &HF0
End Enum

Public Enum CustomLineCapType
    CustomLineCapTypeDefault = 0
    CustomLineCapTypeAdjustableArrow = 1
End Enum

Public Enum GpLineJoin
    LineJoinMiter = 0
    LineJoinBevel = 1
    LineJoinRound = 2
    LineJoinMiterClipped = 3
End Enum

Public Enum PenType
    PenTypeSolidColor = BrushTypeSolidColor
    PenTypeHatchFill = BrushTypeHatchFill
    PenTypeTextureFill = BrushTypeTextureFill
    PenTypePathGradient = BrushTypePathGradient
    PenTypeLinearGradient = BrushTypeLinearGradient
    PenTypeUnknown = -1
End Enum
'=================================

Public Enum MetafileType
    MetafileTypeInvalid
    MetafileTypeWmf
    MetafileTypeWmfPlaceable
    MetafileTypeEmf
    MetafileTypeEmfPlusOnly
    MetafileTypeEmfPlusDual
End Enum

Public Enum emfType
    EmfTypeEmfOnly = MetafileTypeEmf
    EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly
    EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual
End Enum

Public Enum ObjectType
    ObjectTypeInvalid
    ObjectTypeBrush
    ObjectTypePen
    ObjectTypePath
    ObjectTypeRegion
    ObjectTypeImage
    ObjectTypeFont
    ObjectTypeStringFormat
    ObjectTypeImageAttributes
    ObjectTypeCustomLineCap
    ObjectTypeGraphics
    ObjectTypeMax = ObjectTypeGraphics
    'ObjectTypeMax = ObjectTypeCustomLineCap
    ObjectTypeMin = ObjectTypeBrush
End Enum

Public Enum MetafileFrameUnit
    MetafileFrameUnitPixel = UnitPixel
    MetafileFrameUnitPoint = UnitPoint
    MetafileFrameUnitInch = UnitInch
    MetafileFrameUnitDocument = UnitDocument
    MetafileFrameUnitMillimeter = UnitMillimeter
    MetafileFrameUnitGdi
End Enum

Public Enum CoordinateSpace
    CoordinateSpaceWorld     ' 0
    CoordinateSpacePage      ' 1
    CoordinateSpaceDevice    ' 2
End Enum

Public Enum EmfPlusRecordType
    WmfRecordTypeSetBkColor = &H10201
    WmfRecordTypeSetBkMode = &H10102
    WmfRecordTypeSetMapMode = &H10103
    WmfRecordTypeSetROP2 = &H10104
    WmfRecordTypeSetRelAbs = &H10105
    WmfRecordTypeSetPolyFillMode = &H10106
    WmfRecordTypeSetStretchBltMode = &H10107
    WmfRecordTypeSetTextCharExtra = &H10108
    WmfRecordTypeSetTextColor = &H10209
    WmfRecordTypeSetTextJustification = &H1020A
    WmfRecordTypeSetWindowOrg = &H1020B
    WmfRecordTypeSetWindowExt = &H1020C
    WmfRecordTypeSetViewportOrg = &H1020D
    WmfRecordTypeSetViewportExt = &H1020E
    WmfRecordTypeOffsetWindowOrg = &H1020F
    WmfRecordTypeScaleWindowExt = &H10410
    WmfRecordTypeOffsetViewportOrg = &H10211
    WmfRecordTypeScaleViewportExt = &H10412
    WmfRecordTypeLineTo = &H10213
    WmfRecordTypeMoveTo = &H10214
    WmfRecordTypeExcludeClipRect = &H10415
    WmfRecordTypeIntersectClipRect = &H10416
    WmfRecordTypeArc = &H10817
    WmfRecordTypeEllipse = &H10418
    WmfRecordTypeFloodFill = &H10419
    WmfRecordTypePie = &H1081A
    WmfRecordTypeRectangle = &H1041B
    WmfRecordTypeRoundRect = &H1061C
    WmfRecordTypePatBlt = &H1061D
    WmfRecordTypeSaveDC = &H1001E
    WmfRecordTypeSetPixel = &H1041F
    WmfRecordTypeOffsetClipRgn = &H10220
    WmfRecordTypeTextOut = &H10521
    WmfRecordTypeBitBlt = &H10922
    WmfRecordTypeStretchBlt = &H10B23
    WmfRecordTypePolygon = &H10324
    WmfRecordTypePolyline = &H10325
    WmfRecordTypeEscape = &H10626
    WmfRecordTypeRestoreDC = &H10127
    WmfRecordTypeFillRegion = &H10228
    WmfRecordTypeFrameRegion = &H10429
    WmfRecordTypeInvertRegion = &H1012A
    WmfRecordTypePaintRegion = &H1012B
    WmfRecordTypeSelectClipRegion = &H1012C
    WmfRecordTypeSelectObject = &H1012D
    WmfRecordTypeSetTextAlign = &H1012E
    WmfRecordTypeDrawText = &H1062F
    WmfRecordTypeChord = &H10830
    WmfRecordTypeSetMapperFlags = &H10231
    WmfRecordTypeExtTextOut = &H10A32
    WmfRecordTypeSetDIBToDev = &H10D33
    WmfRecordTypeSelectPalette = &H10234
    WmfRecordTypeRealizePalette = &H10035
    WmfRecordTypeAnimatePalette = &H10436
    WmfRecordTypeSetPalEntries = &H10037
    WmfRecordTypePolyPolygon = &H10538
    WmfRecordTypeResizePalette = &H10139
    WmfRecordTypeDIBBitBlt = &H10940
    WmfRecordTypeDIBStretchBlt = &H10B41
    WmfRecordTypeDIBCreatePatternBrush = &H10142
    WmfRecordTypeStretchDIB = &H10F43
    WmfRecordTypeExtFloodFill = &H10548
    WmfRecordTypeSetLayout = &H10149
    WmfRecordTypeResetDC = &H1014C
    WmfRecordTypeStartDoc = &H1014D
    WmfRecordTypeStartPage = &H1004F
    WmfRecordTypeEndPage = &H10050
    WmfRecordTypeAbortDoc = &H10052
    WmfRecordTypeEndDoc = &H1005E
    WmfRecordTypeDeleteObject = &H101F0
    WmfRecordTypeCreatePalette = &H100F7
    WmfRecordTypeCreateBrush = &H100F8
    WmfRecordTypeCreatePatternBrush = &H101F9
    WmfRecordTypeCreatePenIndirect = &H102FA
    WmfRecordTypeCreateFontIndirect = &H102FB
    WmfRecordTypeCreateBrushIndirect = &H102FC
    WmfRecordTypeCreateBitmapIndirect = &H102FD
    WmfRecordTypeCreateBitmap = &H106FE
    WmfRecordTypeCreateRegion = &H106FF
    EmfRecordTypeHeader = 1
    EmfRecordTypePolyBezier = 2
    EmfRecordTypePolygon = 3
    EmfRecordTypePolyline = 4
    EmfRecordTypePolyBezierTo = 5
    EmfRecordTypePolyLineTo = 6
    EmfRecordTypePolyPolyline = 7
    EmfRecordTypePolyPolygon = 8
    EmfRecordTypeSetWindowExtEx = 9
    EmfRecordTypeSetWindowOrgEx = 10
    EmfRecordTypeSetViewportExtEx = 11
    EmfRecordTypeSetViewportOrgEx = 12
    EmfRecordTypeSetBrushOrgEx = 13
    EmfRecordTypeEOF = 14
    EmfRecordTypeSetPixelV = 15
    EmfRecordTypeSetMapperFlags = 16
    EmfRecordTypeSetMapMode = 17
    EmfRecordTypeSetBkMode = 18
    EmfRecordTypeSetPolyFillMode = 19
    EmfRecordTypeSetROP2 = 20
    EmfRecordTypeSetStretchBltMode = 21
    EmfRecordTypeSetTextAlign = 22
    EmfRecordTypeSetColorAdjustment = 23
    EmfRecordTypeSetTextColor = 24
    EmfRecordTypeSetBkColor = 25
    EmfRecordTypeOffsetClipRgn = 26
    EmfRecordTypeMoveToEx = 27
    EmfRecordTypeSetMetaRgn = 28
    EmfRecordTypeExcludeClipRect = 29
    EmfRecordTypeIntersectClipRect = 30
    EmfRecordTypeScaleViewportExtEx = 31
    EmfRecordTypeScaleWindowExtEx = 32
    EmfRecordTypeSaveDC = 33
    EmfRecordTypeRestoreDC = 34
    EmfRecordTypeSetWorldTransform = 35
    EmfRecordTypeModifyWorldTransform = 36
    EmfRecordTypeSelectObject = 37
    EmfRecordTypeCreatePen = 38
    EmfRecordTypeCreateBrushIndirect = 39
    EmfRecordTypeDeleteObject = 40
    EmfRecordTypeAngleArc = 41
    EmfRecordTypeEllipse = 42
    EmfRecordTypeRectangle = 43
    EmfRecordTypeRoundRect = 44
    EmfRecordTypeArc = 45
    EmfRecordTypeChord = 46
    EmfRecordTypePie = 47
    EmfRecordTypeSelectPalette = 48
    EmfRecordTypeCreatePalette = 49
    EmfRecordTypeSetPaletteEntries = 50
    EmfRecordTypeResizePalette = 51
    EmfRecordTypeRealizePalette = 52
    EmfRecordTypeExtFloodFill = 53
    EmfRecordTypeLineTo = 54
    EmfRecordTypeArcTo = 55
    EmfRecordTypePolyDraw = 56
    EmfRecordTypeSetArcDirection = 57
    EmfRecordTypeSetMiterLimit = 58
    EmfRecordTypeBeginPath = 59
    EmfRecordTypeEndPath = 60
    EmfRecordTypeCloseFigure = 61
    EmfRecordTypeFillPath = 62
    EmfRecordTypeStrokeAndFillPath = 63
    EmfRecordTypeStrokePath = 64
    EmfRecordTypeFlattenPath = 65
    EmfRecordTypeWidenPath = 66
    EmfRecordTypeSelectClipPath = 67
    EmfRecordTypeAbortPath = 68
    EmfRecordTypeReserved_069 = 69
    EmfRecordTypeGdiComment = 70
    EmfRecordTypeFillRgn = 71
    EmfRecordTypeFrameRgn = 72
    EmfRecordTypeInvertRgn = 73
    EmfRecordTypePaintRgn = 74
    EmfRecordTypeExtSelectClipRgn = 75
    EmfRecordTypeBitBlt = 76
    EmfRecordTypeStretchBlt = 77
    EmfRecordTypeMaskBlt = 78
    EmfRecordTypePlgBlt = 79
    EmfRecordTypeSetDIBitsToDevice = 80
    EmfRecordTypeStretchDIBits = 81
    EmfRecordTypeExtCreateFontIndirect = 82
    EmfRecordTypeExtTextOutA = 83
    EmfRecordTypeExtTextOutW = 84
    EmfRecordTypePolyBezier16 = 85
    EmfRecordTypePolygon16 = 86
    EmfRecordTypePolyline16 = 87
    EmfRecordTypePolyBezierTo16 = 88
    EmfRecordTypePolylineTo16 = 89
    EmfRecordTypePolyPolyline16 = 90
    EmfRecordTypePolyPolygon16 = 91
    EmfRecordTypePolyDraw16 = 92
    EmfRecordTypeCreateMonoBrush = 93
    EmfRecordTypeCreateDIBPatternBrushPt = 94
    EmfRecordTypeExtCreatePen = 95
    EmfRecordTypePolyTextOutA = 96
    EmfRecordTypePolyTextOutW = 97
    EmfRecordTypeSetICMMode = 98
    EmfRecordTypeCreateColorSpace = 99
    EmfRecordTypeSetColorSpace = 100
    EmfRecordTypeDeleteColorSpace = 101
    EmfRecordTypeGLSRecord = 102
    EmfRecordTypeGLSBoundedRecord = 103
    EmfRecordTypePixelFormat = 104
    EmfRecordTypeDrawEscape = 105
    EmfRecordTypeExtEscape = 106
    EmfRecordTypeStartDoc = 107
    EmfRecordTypeSmallTextOut = 108
    EmfRecordTypeForceUFIMapping = 109
    EmfRecordTypeNamedEscape = 110
    EmfRecordTypeColorCorrectPalette = 111
    EmfRecordTypeSetICMProfileA = 112
    EmfRecordTypeSetICMProfileW = 113
    EmfRecordTypeAlphaBlend = 114
    EmfRecordTypeSetLayout = 115
    EmfRecordTypeTransparentBlt = 116
    EmfRecordTypeReserved_117 = 117
    EmfRecordTypeGradientFill = 118
    EmfRecordTypeSetLinkedUFIs = 119
    EmfRecordTypeSetTextJustification = 120
    EmfRecordTypeColorMatchToTargetW = 121
    EmfRecordTypeCreateColorSpaceW = 122
    EmfRecordTypeMax = 122
    EmfRecordTypeMin = 1
    EmfPlusRecordTypeInvalid = 16384
    EmfPlusRecordTypeHeader = 16385
    EmfPlusRecordTypeEndOfFile = 16386
    EmfPlusRecordTypeComment = 16387
    EmfPlusRecordTypeGetDC = 16388
    EmfPlusRecordTypeMultiFormatStart = 16389
    EmfPlusRecordTypeMultiFormatSection = 16390
    EmfPlusRecordTypeMultiFormatEnd = 16391
    EmfPlusRecordTypeObject = 16392
    EmfPlusRecordTypeClear = 16393
    EmfPlusRecordTypeFillRects = 16394
    EmfPlusRecordTypeDrawRects = 16395
    EmfPlusRecordTypeFillPolygon = 16396
    EmfPlusRecordTypeDrawLines = 16397
    EmfPlusRecordTypeFillEllipse = 16398
    EmfPlusRecordTypeDrawEllipse = 16399
    EmfPlusRecordTypeFillPie = 16400
    EmfPlusRecordTypeDrawPie = 16401
    EmfPlusRecordTypeDrawArc = 16402
    EmfPlusRecordTypeFillRegion = 16403
    EmfPlusRecordTypeFillPath = 16404
    EmfPlusRecordTypeDrawPath = 16405
    EmfPlusRecordTypeFillClosedCurve = 16406
    EmfPlusRecordTypeDrawClosedCurve = 16407
    EmfPlusRecordTypeDrawCurve = 16408
    EmfPlusRecordTypeDrawBeziers = 16409
    EmfPlusRecordTypeDrawImage = 16410
    EmfPlusRecordTypeDrawImagePoints = 16411
    EmfPlusRecordTypeDrawString = 16412
    EmfPlusRecordTypeSetRenderingOrigin = 16413
    EmfPlusRecordTypeSetAntiAliasMode = 16414
    EmfPlusRecordTypeSetTextRenderingHint = 16415
    EmfPlusRecordTypeSetTextContrast = 16416
    EmfPlusRecordTypeSetInterpolationMode = 16417
    EmfPlusRecordTypeSetPixelOffsetMode = 16418
    EmfPlusRecordTypeSetCompositingMode = 16419
    EmfPlusRecordTypeSetCompositingQuality = 16420
    EmfPlusRecordTypeSave = 16421
    EmfPlusRecordTypeRestore = 16422
    EmfPlusRecordTypeBeginContainer = 16423
    EmfPlusRecordTypeBeginContainerNoParams = 16424
    EmfPlusRecordTypeEndContainer = 16425
    EmfPlusRecordTypeSetWorldTransform = 16426
    EmfPlusRecordTypeResetWorldTransform = 16427
    EmfPlusRecordTypeMultiplyWorldTransform = 16428
    EmfPlusRecordTypeTranslateWorldTransform = 16429
    EmfPlusRecordTypeScaleWorldTransform = 16430
    EmfPlusRecordTypeRotateWorldTransform = 16431
    EmfPlusRecordTypeSetPageTransform = 16432
    EmfPlusRecordTypeResetClip = 16433
    EmfPlusRecordTypeSetClipRect = 16434
    EmfPlusRecordTypeSetClipPath = 16435
    EmfPlusRecordTypeSetClipRegion = 16436
    EmfPlusRecordTypeOffsetClip = 16437
    EmfPlusRecordTypeDrawDriverString = 16438
    EmfPlusRecordTypeStrokeFillPath = 16439
    EmfPlusRecordTypeSerializableObject = 16440
    EmfPlusRecordTypeSetTSGraphics = 16441
    EmfPlusRecordTypeSetTSClip = 16442
    EmfPlusRecordTotal = 16443
    'EmfPlusRecordTotal = 16439
    EmfPlusRecordTypeMax = EmfPlusRecordTotal - 1
    EmfPlusRecordTypeMin = EmfPlusRecordTypeHeader
End Enum
'=================================

Public Enum GpHotkeyPrefix
    HotkeyPrefixNone = 0
    HotkeyPrefixShow = 1
    HotkeyPrefixHide = 2
End Enum

Public Enum FlushIntention
    FlushIntentionFlush = 0
    FlushIntentionSync = 1
End Enum

Public Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6
    EncoderParameterValueTypeUndefined = 7
    EncoderParameterValueTypeRationalRange = 8
    EncoderParameterValueTypePointer = 9
End Enum

Public Enum EncoderValue
    EncoderValueColorTypeCMYK = 0
    EncoderValueColorTypeYCCK
    EncoderValueCompressionLZW
    EncoderValueCompressionCCITT3
    EncoderValueCompressionCCITT4
    EncoderValueCompressionRle
    EncoderValueCompressionNone
    EncoderValueScanMethodInterlaced
    EncoderValueScanMethodNonInterlaced
    EncoderValueVersionGif87
    EncoderValueVersionGif89
    EncoderValueRenderProgressive
    EncoderValueRenderNonProgressive
    EncoderValueTransformRotate90
    EncoderValueTransformRotate180
    EncoderValueTransformRotate270
    EncoderValueTransformFlipHorizontal
    EncoderValueTransformFlipVertical
    EncoderValueMultiFrame
    EncoderValueLastFrame
    EncoderValueFlush
    EncoderValueFrameDimensionTime
    EncoderValueFrameDimensionResolution
    EncoderValueFrameDimensionPage
    EncoderValueColorTypeGray
    EncoderValueColorTypeRGB
End Enum

Public Enum ConvertToEmfPlusFlags
    ConvertToEmfPlusFlagsDefault = 0
    ConvertToEmfPlusFlagsRopUsed = 1
    ConvertToEmfPlusFlagsText = 2
    ConvertToEmfPlusFlagsInvalidRecord = 4
End Enum

Public Enum DebugEventLevel
    DebugEventLevelFatal = 0
    DebugEventLevelWarning
End Enum

Public Enum GpTestControlEnum
    TestControlForceBilinear = 0
    TestControlNoICM = 1
    TestControlGetBuildNumber = 2
End Enum

Public Enum GpStatus
    Ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

Public Enum GdipImageType
    Bmp
    EMF
    WMF
    JPG
    PNG
    GIF
    TIF
    ICO
End Enum

Public Enum BoolType
    FALSEVALUE = 0
    TRUEVALUE = 1
End Enum

Public Enum DriverStringOptions
  DriverStringOptionsCmapLookup
  DriverStringOptionsVertical
  DriverStringOptionsRealizedAdvance
  DriverStringOptionsLimitSubpixel
End Enum