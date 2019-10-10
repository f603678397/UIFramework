VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWidgetManager"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum PresetThemeEnum
    DrakTheme
    LightTheme
End Enum

Public Sub Initialize(Optional ByVal nFlag As PresetThemeEnum)
    SetPresetTheme nFlag
End Sub

Public Sub SetPresetTheme(ByVal nFlag As PresetThemeEnum)
    Select Case nFlag
        Case PresetThemeEnum.DrakTheme
            SetDarkTheme
        Case PresetThemeEnum.LightTheme
            SetLightTheme
    End Select
End Sub

Private Sub SetDarkTheme()
    With ThemeColor
        .AccentColor = cColor.FromARGB(255, 30, 110, 195)
        .AccentDarkColor = cColor.SetBrightness(.AccentColor, -0.1)
        .AccentLightColor = cColor.SetBrightness(.AccentColor, 0.1)
        .PrimaryColor = cColor.FromARGB(255, 70, 70, 70)
        .PrimaryDrakColor = cColor.FromARGB(255, 65, 65, 65)
        .PrimaryLightColor = cColor.FromARGB(255, 80, 80, 80)
        .BorderColor = cColor.FromARGB(255, 40, 40, 40)
        .TextColor = cColor.FromARGB(255, 205, 205, 205)
    End With
End Sub

Private Sub SetLightTheme()
    With ThemeColor
        .AccentColor = cColor.FromARGB(255, 50, 210, 130)
        .AccentDarkColor = cColor.SetBrightness(.AccentColor, -0.1)
        .AccentLightColor = cColor.SetBrightness(.AccentColor, 0.1)
        .PrimaryColor = cColor.FromARGB(255, 240, 240, 240)
        .PrimaryDrakColor = cColor.FromARGB(255, 235, 235, 235)
        .PrimaryLightColor = cColor.FromARGB(255, 250, 250, 250)
        .BorderColor = cColor.FromARGB(255, 200, 200, 200)
        .TextColor = cColor.FromARGB(255, 50, 50, 50)
    End With
End Sub

