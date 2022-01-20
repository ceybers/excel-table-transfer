Attribute VB_Name = "modImageListHelper"
'@Folder("HelperFunctions")
Option Explicit

Public Sub PopulateImageList(ByRef il As ImageList, ByVal iconSize As Integer)
    Debug.Assert Not il Is Nothing
    AddImageListImage il, "root", "BlogHomePage", iconSize
    AddImageListImage il, "wb", "FileSaveAsExcelXlsx", iconSize
    AddImageListImage il, "ws", "HeaderFooterSheetNameInsert", iconSize
    AddImageListImage il, "lo", "CreateTable", iconSize
    AddImageListImage il, "col", "TableColumnSelect", iconSize
    AddImageListImage il, "activeLo", "TableAutoFormatStyle", iconSize
    AddImageListImage il, "delete", "Delete", iconSize
    AddImageListImage il, "FindText", "FindText", iconSize
    AddImageListImage il, "AdpPrimaryKey", "AdpPrimaryKey", iconSize
End Sub

Private Sub AddImageListImage(ByRef il As ImageList, ByVal key As String, ByVal imageMso As String, ByVal iconSize As Integer)
    il.ListImages.Add , key, Application.CommandBars.GetImageMso(imageMso, iconSize, iconSize)
End Sub

