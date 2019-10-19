Attribute VB_Name = "mGlobal"
Option Explicit

Public Type ThemeColorType
    PrimaryColor        As Long
    PrimaryDrakColor    As Long
    PrimaryLightColor   As Long
    BorderColor         As Long
    TextColor           As Long
    AccentColor         As Long
    AccentDarkColor     As Long
    AccentLightColor    As Long
End Type

Public ViewList         As New cObjectList

Public Preset           As PresetThemeEnum

Public Sub SetDarkTheme(ByRef ThemeColor As ThemeColorType)
    With ThemeColor
        .AccentColor = cColor.FromARGB(255, 30, 110, 195)
        .AccentDarkColor = cColor.SetBrightness(.AccentColor, -0.1)
        .AccentLightColor = cColor.SetBrightness(.AccentColor, 0.15)
        .PrimaryColor = cColor.FromARGB(255, 70, 70, 70)
        .PrimaryDrakColor = cColor.FromARGB(255, 65, 65, 65)
        .PrimaryLightColor = cColor.FromARGB(255, 85, 85, 85)
        .BorderColor = cColor.FromARGB(255, 110, 110, 110)
        .TextColor = cColor.FromARGB(255, 205, 205, 205)
    End With
End Sub

Public Sub SetLightTheme(ByRef ThemeColor As ThemeColorType)
    With ThemeColor
        .AccentColor = cColor.FromARGB(255, 45, 175, 70)
        .AccentDarkColor = cColor.SetBrightness(.AccentColor, -0.1)
        .AccentLightColor = cColor.SetBrightness(.AccentColor, 0.15)
        .PrimaryColor = cColor.FromARGB(255, 240, 240, 240)
        .PrimaryDrakColor = cColor.FromARGB(255, 235, 235, 235)
        .PrimaryLightColor = cColor.FromARGB(255, 250, 250, 250)
        .BorderColor = cColor.FromARGB(255, 200, 200, 200)
        .TextColor = cColor.FromARGB(255, 50, 50, 50)
    End With
End Sub
