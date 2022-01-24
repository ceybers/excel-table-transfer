Attribute VB_Name = "modStandardImageList"
'@Folder "HelperFunctions"
Option Explicit

Public Function GetMSOImageList(Optional ByVal iconSize As Integer = 16) As ImageList
    Set GetMSOImageList = New ImageList

    AddImageListImage GetMSOImageList, "root", "BlogHomePage", iconSize
    AddImageListImage GetMSOImageList, "wb", "FileSaveAsExcelXlsx", iconSize
    AddImageListImage GetMSOImageList, "ws", "HeaderFooterSheetNameInsert", iconSize
    AddImageListImage GetMSOImageList, "lo", "CreateTable", iconSize
    AddImageListImage GetMSOImageList, "col", "TableColumnSelect", iconSize
    AddImageListImage GetMSOImageList, "activeLo", "TableAutoFormatStyle", iconSize
    AddImageListImage GetMSOImageList, "delete", "Delete", iconSize
    AddImageListImage GetMSOImageList, "AutoSum", "AutoSum", iconSize
    AddImageListImage GetMSOImageList, "MagicWand", "QueryBuilder", iconSize
    AddImageListImage GetMSOImageList, "TraceError", "TraceError", iconSize
    AddImageListImage GetMSOImageList, "AcceptInvitation", "AcceptInvitation", iconSize
    AddImageListImage GetMSOImageList, "DeclineInvitation", "DeclineInvitation", iconSize
End Function

Private Sub AddImageListImage(ByRef il As ImageList, ByVal key As String, ByVal imageMso As String, ByVal iconSize As Integer)
    il.ListImages.Add 1, key, Application.CommandBars.GetImageMso(imageMso, iconSize, iconSize)
End Sub
