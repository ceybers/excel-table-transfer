Attribute VB_Name = "modLoadImageLists"
'@Folder("MVVM.Infrastructure")
Option Explicit

Private Const MSO_LIST As String = "Column,SelectTaskColumn;Yes,WorkflowComplete;No,CancelRequest;" & _
    "TypeText,DataTypeText;TypeNumber,DataTypeNumber"

Public Sub LoadImageLists(ByVal ImageLists As Scripting.Dictionary)
    LoadImageList ImageLists, 16
    LoadImageList ImageLists, 24
    LoadImageList ImageLists, 32
    LoadImageList ImageLists, 48
    LoadImageList ImageLists, 64
End Sub

Private Sub LoadImageList(ByVal ImageLists As Scripting.Dictionary, ByVal PictureSize As Long)
    Dim Result As ImageList
    Set Result = New ImageList
    
    LoadImagesToImageList Result, PictureSize
    
    ImageLists.Add Key:=CStr(PictureSize), Item:=Result
End Sub

Private Sub LoadImagesToImageList(ByVal ImageList As ImageList, ByVal PictureSize As Long)
    Dim Element As Variant
    For Each Element In Split(MSO_LIST, ";")
        Dim Key As String
        Dim Value As String
        Key = Split(Element, ",")(0)
        Value = Split(Element, ",")(1)
        ImageList.ListImages.Add Key:=Key, _
            Picture:=Application.CommandBars.GetImageMso(Value, PictureSize, PictureSize)
    Next Element
End Sub

