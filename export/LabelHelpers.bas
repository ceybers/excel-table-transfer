Attribute VB_Name = "LabelHelpers"
'@Folder("Helpers")
Option Explicit

Public Sub ApplyImageMSOtoLabel(ByVal Label As Object, ByVal ImageMSOKey As String)
    Dim Picture As StdPicture
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set Picture = Application.CommandBars.GetImageMso(ImageMSOKey, 32, 32)
    Set Label.Picture = Picture
End Sub
