Attribute VB_Name = "StandardImageList"
'@Folder "Helpers.Controls"
Option Explicit

Private Const DEFAULT_ICON_SIZE As Long = 16

'@Description "Returns a new ImageList object pre-populated with a standardised list of default icons."
Public Function GetMSOImageList(Optional ByVal IconSize As Long = DEFAULT_ICON_SIZE) As ImageList
Attribute GetMSOImageList.VB_Description = "Returns a new ImageList object pre-populated with a standardised list of default icons."
    Dim Result As ImageList
    Set Result = New ImageList
    
    Dim Controls As Controls
    Set Controls = frmPictures16.Controls
    If IconSize = 32 Then Set Controls = frmPictures32.Controls
    
    Dim Control As Control
    For Each Control In Controls
        If TypeOf Control Is MSForms.Label Then
            Dim Label As MSForms.Label
            Set Label = Control
            Result.ListImages.Add Key:=Label.Name, Picture:=Label.Picture
        End If
    Next Control
    
    Set GetMSOImageList = Result
End Function

Private Sub AddImageToImageList(ByVal ImageList As ImageList, ByVal Key As String, ByVal ImageMso As String, ByVal IconSize As Long)
    Dim Picture As IPictureDisp
    Set Picture = Application.CommandBars.GetImageMso(ImageMso, IconSize, IconSize)
    ImageList.ListImages.Add 1, Key, Picture
End Sub
