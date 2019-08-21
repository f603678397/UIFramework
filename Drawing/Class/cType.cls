VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cType"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'??
'----------------------------------
Public Type POINTL
    X As Long
    Y As Long
End Type
'----------------------------------
Public Type POINTF
    X As Single
    Y As Single
End Type
'==================================

'????
'----------------------------------
Public Type RECTL
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
'----------------------------------
Public Type RECTF
    Left   As Single
    Top    As Single
    Right  As Single
    Bottom As Single
End Type
'==================================

'???
'----------------------------------
Public Type SIZEL
    cx As Long
    cy As Long
End Type
'----------------------------------
Public Type SIZEF
    cx As Single
    cy As Single
End Type
'==================================

'???
'----------------------------------
Public Type RGBQUAD
    rgbBlue     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type
'----------------------------------
Public Type COLORBYTES
    BlueByte  As Byte
    GreenByte As Byte
    RedByte   As Byte
    AlphaByte As Byte
End Type
'----------------------------------
Public Type COLORLONG
    longval As Long
End Type

Public Type ColorMap
    oldColor As Long
    newColor As Long
End Type
'----------------------------------
Public Type ColorMatrix
    M(0 To 4, 0 To 4) As Single
End Type
'==================================

'¦Ë?
'----------------------------------
Public Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
'----------------------------------
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
'----------------------------------
Public Type BitmapData
    Width       As Long
    Height      As Long
    stride      As Long
    PixelFormat As GpPixelFormat
    Scan0       As Long
    Reserved    As Long
End Type
'==================================

'¡¤??
'----------------------------------
Public Type PathData
    Count   As Long
    pPoints As Long
    pTypes  As Long
End Type
'==================================

'????
'----------------------------------
Public Type CLSID
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type
'----------------------------------
Public Type EncoderParameter
    GUID           As CLSID
    NumberOfValues As Long

    Type           As EncoderParameterValueType
    value          As Long
End Type
'----------------------------------
Public Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type
'==================================
Public Type GdiplusAbort
    AbortCallback As Long
End Type

Public Type ImageItemData
    Size     As Long
    Position As Long
    pDesc    As Long
    DescSize As Long
    pData    As Long
    dataSize As Long
    Cookie   As Long
End Type

Public Type SharpenParams
    radius As Single
    amount As Single
End Type

Public Type BlurParams
    radius     As Single
    expandEdge As Long
End Type

Public Type BrightnessContrastParams
    brightnessLevel As Long
    contrastLevel   As Long
End Type

Public Type RedEyeCorrectionParams
    numberOfAreas As Long
    areas         As RECTL
End Type

Public Type HueSaturationLightnessParams
    hueLevel        As Long
    saturationLevel As Long
    lightnessLevel  As Long
End Type

Public Type TintParams
    Hue    As Long
    amount As Long
End Type

Public Type LevelsParams
    highlight As Long
    midtone   As Long
    Shadow    As Long
End Type

Public Type ColorBalanceParams
    cyanRed      As Long
    magentaGreen As Long
    yellowBlue   As Long
End Type

Public Type ColorLUTParams
    lutB(0 To 255) As Byte
    lutG(0 To 255) As Byte
    lutR(0 To 255) As Byte
    lutA(0 To 255) As Byte
End Type

Public Type ColorCurveParams
    adjustment  As CurveAdjustments
    channel     As CurveChannel
    adjustValue As Long
End Type

Public Type LOGFONTA
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type

Public Type LOGFONTW
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type

Public Type ImageCodecInfo
    ClassID           As CLSID
    FormatID          As CLSID
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As ImageCodecFlags
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type

Public Type ColorPalette
    Flags             As PaletteFlags
    Count             As Long
    Entries(0 To 255) As Long
End Type

Public Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Public Type WmfPlaceableFileHeader
    Key         As Long
    Hmf         As Integer
    boundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    Checksum    As Integer
End Type

Public Type ENHMETAHEADER3
    itype          As Long
    nSize          As Long
    rclBounds      As RECTL
    rclFrame       As RECTL
    dSignature     As Long
    nVersion       As Long
    nBytes         As Long
    nRecords       As Long
    nHandles       As Integer
    sReserved      As Integer
    nDescription   As Long
    offDescription As Long
    nPalEntries    As Long
    szlDevice      As SIZEL
    szlMillimeters As SIZEL
End Type

Public Type METAHEADER
    mtType         As Integer
    mtHeaderSize   As Integer
    mtVersion      As Integer
    mtSize         As Long
    mtNoObjects    As Integer
    mtMaxRecord    As Long
    mtNoParameters As Integer
End Type

Public Type MetafileHeader
    mType             As MetafileType
    Size              As Long
    Version           As Long
    EmfPlusFlags      As Long
    DpiX              As Single
    DpiY              As Single
    X                 As Long
    Y                 As Long
    Width             As Long
    Height            As Long
    EmfHeader         As ENHMETAHEADER3
    EmfPlusHeaderSize As Long
    LogicalDpiX       As Long
    LogicalDpiY       As Long
End Type

Public Type PropertyItem
    propId As Long
    Length As Long
    Type   As Integer
    value  As Long
End Type

Public Type CharacterRange
    first  As Long
    Length As Long
End Type

Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type