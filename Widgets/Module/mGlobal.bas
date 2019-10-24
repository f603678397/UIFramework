Attribute VB_Name = "mGlobal"
Option Explicit

Public Type ThemeColorType
    BKColor                 As Long
    BKDrakColor             As Long
    BKLightColor            As Long
    BorderColor             As Long
    BorderDisEnableColor    As Long
    TextColor               As Long
    AccentColor             As Long
    AccentDarkColor         As Long
    AccentLightColor        As Long
End Type

Public ViewList         As New cObjectList

Public Preset           As PresetThemeEnum

Public Sub SetDarkTheme(ByRef ThemeColor As ThemeColorType)
    With ThemeColor
        .AccentColor = cColor.FromARGB(255, 30, 110, 195)
        .AccentDarkColor = cColor.SetBrightness(.AccentColor, -0.2)
        .AccentLightColor = cColor.SetBrightness(.AccentColor, 0.15)
        .BKColor = cColor.FromARGB(255, 70, 70, 70)
        .BKDrakColor = cColor.FromARGB(255, 65, 65, 65)
        .BKLightColor = cColor.FromARGB(255, 85, 85, 85)
        .BorderColor = cColor.FromARGB(255, 110, 110, 110)
        .BorderDisEnableColor = cColor.FromARGB(255, 80, 80, 80)
        .TextColor = cColor.FromARGB(255, 205, 205, 205)
    End With
End Sub

Public Sub SetLightTheme(ByRef ThemeColor As ThemeColorType)
    With ThemeColor
        .AccentColor = cColor.FromARGB(255, 45, 175, 70)
        .AccentDarkColor = cColor.SetBrightness(.AccentColor, -0.2)
        .AccentLightColor = cColor.SetBrightness(.AccentColor, 0.15)
        .BKColor = cColor.FromARGB(255, 248, 248, 248)
        .BKDrakColor = cColor.FromARGB(255, 240, 240, 240)
        .BKLightColor = cColor.FromARGB(255, 255, 255, 255)
        .BorderColor = cColor.FromARGB(255, 200, 200, 200)
        .BorderDisEnableColor = cColor.FromARGB(255, 205, 205, 205)
        .TextColor = cColor.FromARGB(255, 50, 50, 50)
    End With
End Sub
