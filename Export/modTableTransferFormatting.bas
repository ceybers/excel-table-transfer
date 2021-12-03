Attribute VB_Name = "modTableTransferFormatting"
Option Explicit
    
Sub FormatMatch(rng As Range)
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rng.Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.499984740745262
    End With
End Sub

Sub FormatNonMatch(rng As Range)
    With rng.Interior
        .Pattern = xlLightUp
        .PatternColor = 16751103
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub FormatReset(rng As Range)
    With rng.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rng.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub

