VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mHandle As Long

Private Sub Class_Terminate(): Release: End Sub

Public Property Get Handle() As Long: Handle = mHandle: End Property
Friend Property Let Handle(ByVal nHandle As Long): mHandle = nHandle: End Property

Public Sub Create(nType As GdipEffectType): Release: GdipCreateEffect2 nType, mHandle: End Sub
Friend Sub Release()
    If mHandle = 0 Or Token = 0 Then Exit Sub
    GdipDeleteEffect mHandle
    mHandle = 0
End Sub

Public Property Get ParameterSize() As Long: GdipGetEffectParameterSize mHandle, ParameterSize: End Property

Public Sub SetBlur(Blur As BlurParams)
    GdipSetEffectParameters mHandle, Blur, LenB(Blur)
End Sub

Public Sub SetBrightness(Brightness As BrightnessContrastParams)
    GdipSetEffectParameters mHandle, Brightness, LenB(Brightness)
End Sub

Public Sub SetColorBalance(Balance As ColorBalanceParams)
    GdipSetEffectParameters mHandle, Balance, LenB(Balance)
End Sub

Public Sub SetColorCurve(Curve As ColorCurveParams)
    GdipSetEffectParameters mHandle, Curve, LenB(Curve)
End Sub

Public Sub SetColorLookupTable(LUT As ColorLUTParams)
    GdipSetEffectParameters mHandle, LUT, LenB(LUT)
End Sub

Public Sub SetColorMatrix(M As ColorMatrix)
    GdipSetEffectParameters mHandle, M, LenB(M)
End Sub

Public Sub SetHueSaturationLightness(Hue As HueSaturationLightnessParams)
    GdipSetEffectParameters mHandle, Hue, LenB(Hue)
End Sub

Public Sub SetLevels(Level As LevelsParams)
    GdipSetEffectParameters mHandle, Level, LenB(Level)
End Sub

Public Sub SetRedEyeCorrection(RedEye As RedEyeCorrectionParams)
    GdipSetEffectParameters mHandle, RedEye, LenB(RedEye)
End Sub

Public Sub SetSharpen(S As SharpenParams)
    GdipSetEffectParameters mHandle, S, LenB(S)
End Sub

Public Sub SetTint(T As TintParams)
    GdipSetEffectParameters mHandle, T, LenB(T)
End Sub
