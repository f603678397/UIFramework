Attribute VB_Name = "mFunction"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As CLSID) As Long

Public Function GetErrorDescription(ByVal ErrorNumber As Long) As String
    Dim Description As String
    
    Select Case ErrorNumber
        Case GpStatus.Ok
            Description = "???"
        Case GpStatus.GenericError
            Description = "????????"
        Case GpStatus.InvalidParameter
            Description = "?????????"
        Case GpStatus.OutOfMemory
            Description = "??????"
        Case GpStatus.ObjectBusy
            Description = "????¶ƒ???"
        Case GpStatus.InsufficientBuffer
            Description = "???÷Œ??"
        Case GpStatus.NotImplemented
            Description = "????¶ƒ???"
        Case GpStatus.Win32Error
            Description = "??????"
        Case GpStatus.WrongState
            Description = "??????"
        Case GpStatus.Aborted
            Description = "???????"
        Case GpStatus.FileNotFound
            Description = "???¶ƒ???"
        Case GpStatus.ValueOverflow
            Description = "?????????¶∂"
        Case GpStatus.AccessDenied
            Description = "???????"
        Case GpStatus.UnknownImageFormat
            Description = "¶ƒ????????"
        Case GpStatus.FontFamilyNotFound
            Description = "????¶ƒ???"
        Case GpStatus.FontStyleNotFound
            Description = "???¶ƒ???"
        Case GpStatus.NotTrueTypeFont
            Description = "?????çI??TrueType"
        Case GpStatus.UnsupportedGdiplusVersion
            Description = "??????GDI+?∑⁄"
        Case GpStatus.GdiplusNotInitialized
            Description = "GDI+¶ƒ?????"
        Case GpStatus.PropertyNotFound
            Description = "????¶ƒ???"
        Case GpStatus.PropertyNotSupported
            Description = "????????"
        Case GpStatus.ProfileNotFound
            Description = "???????¶ƒ???"
    End Select
    GetErrorDescription = Description
End Function

Public Sub GdipCreateEffect2(ByVal EffectType As GdipEffectType, Effect As Long)
    Select Case EffectType
        Case GdipEffectType.Blur:                   GdipCreateEffect &H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, Effect
        Case GdipEffectType.BrightnessContrast:     GdipCreateEffect &HD3A1DBE1, &H4C178EC4, &H97EA4C9F, &H3D341CAD, Effect
        Case GdipEffectType.ColorBalance:           GdipCreateEffect &H537E597D, &H48DA251E, &HCA296496, &HF8706B49, Effect
        Case GdipEffectType.ColorCurve:             GdipCreateEffect &HDD6A0022, &H4A6758E4, &H8ED49B9D, &H3DA581B8, Effect
        Case GdipEffectType.ColorLookupTable:       GdipCreateEffect &HA7CE72A9, &H40D70F7F, &HC0D0CCB3, &H12325C2D, Effect
        Case GdipEffectType.ColorMatrix:            GdipCreateEffect &H718F2615, &H40E37933, &H685F11A5, &H74DD14FE, Effect
        Case GdipEffectType.HueSaturationLightness: GdipCreateEffect &H8B2DD6C3, &H4D87EB07, &H871F0A5, &H5F9C6AE2, Effect
        Case GdipEffectType.levels:                 GdipCreateEffect &H99C354EC, &H4F3A2A31, &HA817348C, &H253AB303, Effect
        Case GdipEffectType.RedEyeCorrection:       GdipCreateEffect &H74D29D05, &H426669A4, &HC53C4995, &H32B63628, Effect
        Case GdipEffectType.Sharpen:                GdipCreateEffect &H63CBF3EE, &H402CC526, &HC562718F, &H4251BF40, Effect
        Case GdipEffectType.Tint:                   GdipCreateEffect &H1077AF00, &H44412848, &HAD448994, &H2C7A2D4C, Effect
    End Select
End Sub

Public Function ARGB(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As Long
    Dim nColor As Long
    
    CopyMemory ByVal VarPtr(nColor), B, 1
    CopyMemory ByVal VarPtr(nColor) + 1, G, 1
    CopyMemory ByVal VarPtr(nColor) + 2, R, 1
    CopyMemory ByVal VarPtr(nColor) + 3, A, 1
    
    ARGB = nColor
End Function