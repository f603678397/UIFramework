Attribute VB_Name = "mGlobal"
Option Explicit

Public Type ThemeColorType
    BKColor                 As Long
    BKDrakColor             As Long
    BKLightColor            As Long
    BorderColor             As Long
    BorderDisEnableColor    As Long
    BorderLightColor        As Long
    TextColor               As Long
    AccentColor             As Long
    AccentDarkColor         As Long
    AccentLightColor        As Long
End Type

Public Const PI As Single = 3.1415926

Public ViewList         As New cObjectList

Public Preset           As PresetThemeEnum

Public Sub SetDarkTheme(ByRef ThemeColor As ThemeColorType)
    With ThemeColor
        .AccentColor = cColor.FromARGB(255, 30, 110, 195)
        .AccentDarkColor = cColor.SetBrightness(.AccentColor, -0.2)
        .AccentLightColor = cColor.SetBrightness(.AccentColor, 0.2)
        .BKColor = cColor.FromARGB(255, 70, 70, 70)
        .BKDrakColor = cColor.FromARGB(255, 65, 65, 65)
        .BKLightColor = cColor.FromARGB(255, 85, 85, 85)
        .BorderColor = cColor.FromARGB(255, 110, 110, 110)
        .BorderDisEnableColor = cColor.FromARGB(255, 80, 80, 80)
        .BorderLightColor = cColor.FromARGB(255, 180, 180, 180)
        .TextColor = cColor.FromARGB(255, 205, 205, 205)
    End With
End Sub

Public Sub SetLightTheme(ByRef ThemeColor As ThemeColorType)
    With ThemeColor
        .AccentColor = cColor.FromARGB(255, 45, 175, 70)
        .AccentDarkColor = cColor.SetBrightness(.AccentColor, -0.2)
        .AccentLightColor = cColor.SetBrightness(.AccentColor, 0.2)
        .BKColor = cColor.FromARGB(255, 248, 248, 248)
        .BKDrakColor = cColor.FromARGB(255, 240, 240, 240)
        .BKLightColor = cColor.FromARGB(255, 255, 255, 255)
        .BorderColor = cColor.FromARGB(255, 180, 180, 180)
        .BorderDisEnableColor = cColor.FromARGB(255, 210, 210, 210)
        .BorderLightColor = cColor.FromARGB(255, 140, 140, 140)
        .TextColor = cColor.FromARGB(255, 50, 50, 50)
    End With
End Sub
