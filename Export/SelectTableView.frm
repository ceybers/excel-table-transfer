VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectTableView 
   Caption         =   "Select Table"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "SelectTableView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectTableView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SelectTable"
Option Explicit
Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As SelectTableViewModel
Attribute vm.VB_VarHelpID = -1
Private Const ICON_SIZE As Integer = 16
Private msoImageList As ImageList

Private Type TView
    ' Context as MVVM.IAppContext
    'ViewModel As SelectTableViewModel
    IsCancelled As Boolean
End Type

Private This As TView

Private Sub cmbCancel_Click()
    
    OnCancel
End Sub

Private Sub cmbClearSearch_Click()
    Me.txtSearch = vbNullString
    vm.SearchCriteria = vbNullString
    Me.txtSearch.SetFocus
End Sub

Private Sub cmbOK_Click()
    Me.Hide
End Sub

Private Sub tvTables_DblClick()
    If Not vm.SelectedTable Is Nothing Then
        Me.Hide
    End If
End Sub

Private Sub tvTables_NodeClick(ByVal Node As MSComctlLib.Node)
    vm.TrySelect Node.key
End Sub

Private Sub vm_CollectionChanged()
    LoadTreeview
End Sub

Private Sub vm_ItemSelected()
    If vm.SelectedTable Is Nothing Then
        Me.cmbOK.Enabled = False
    Else
        Me.cmbOK.Enabled = True
    End If
End Sub

Private Sub txtSearch_Change()
    vm.SearchCriteria = txtSearch & "*"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub LoadTreeview()
    With Me.tvTables
        .ImageList = msoImageList
        
        .Nodes.Clear
        .Nodes.Add key:="Root", text:="Excel", image:="root"
        
        .LineStyle = tvwTreeLines
        .Appearance = cc3D
        .Style = tvwTreelinesPlusMinusPictureText
        .Indentation = ICON_SIZE
        .LabelEdit = tvwManual
        .HideSelection = False
    End With
    
    Dim lo As ListObject
    For Each lo In vm.Tables
        TryAddNode lo.parent.parent
        TryAddNode lo.parent
        TryAddNode lo
    Next lo
    
    Dim nd As Node
    For Each nd In Me.tvTables.Nodes
        nd.Expanded = True
    Next nd
End Sub

Private Sub TryHighlightActive()
    Dim nd As Node
    For Each nd In Me.tvTables.Nodes
        If nd.key = vm.ActiveTable.Range.Address(external:=True) Then
            nd.Selected = True
            nd.EnsureVisible
        End If
    Next nd
End Sub

Private Sub TryAddNode(ByVal obj As Object)
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim wb As Workbook
    
    Dim key As String
    Dim parent As String
    Dim nd As Node
    Dim image As String
    Dim text As String
    
    If TypeOf obj Is Workbook Then
        Set wb = obj
        key = "[" & wb.Name & "]"
        parent = "Root"
        image = "wb"
        text = wb.Name
        
    ElseIf TypeOf obj Is Worksheet Then
        Set ws = obj
        key = "[" & ws.parent.Name & "]" & ws.Name
        parent = "[" & ws.parent.Name & "]"
        image = "ws"
        text = ws.Name
        
    ElseIf TypeOf obj Is ListObject Then
        Set lo = obj
        key = lo.Range.Address(external:=True)
        parent = "[" & lo.parent.parent.Name & "]" & lo.parent.Name
        image = "lo"
        text = lo.Name
        If Not vm.ActiveTable Is Nothing Then
            If vm.ActiveTable.Range.Address(external:=True) = lo.Range.Address(external:=True) Then
                image = "activeLo"
                text = text & " (active)"
            End If
        End If
    End If
    
    For Each nd In Me.tvTables.Nodes
        If nd.key = key Then Exit Sub
    Next nd
    
    Me.tvTables.Nodes.Add relative:=parent, Relationship:=tvwChild, key:=key, text:=text, image:=image
End Sub

Private Sub PopulateImageList()
    Dim il As ImageList
    Set msoImageList = New ImageList
    Set il = msoImageList
    
    AddImageListImage il, "root", "BlogHomePage"
    AddImageListImage il, "wb", "FileSaveAsExcelXlsx"
    AddImageListImage il, "ws", "HeaderFooterSheetNameInsert"
    AddImageListImage il, "lo", "CreateTable"
    AddImageListImage il, "col", "TableColumnSelect"
    AddImageListImage il, "activeLo", "TableAutoFormatStyle"
    AddImageListImage il, "delete", "Delete"
End Sub

Private Sub AddImageListImage(ByVal il As ImageList, ByVal key As String, ByVal imageMso As String)
    il.ListImages.Add 1, key, Application.CommandBars.GetImageMso(imageMso, ICON_SIZE, ICON_SIZE)
End Sub

Private Sub InitializeView()
    This.IsCancelled = False
    Me.txtSearch = vbNullString
    Me.cmbOK.Enabled = False
    
    PopulateImageList
    LoadTreeview
    
    Set Me.cmbClearSearch.Picture = msoImageList.ListImages.Item("delete").Picture
    Me.txtSearch.SetFocus
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As IViewModel) As Boolean
    Set vm = ViewModel
    
    InitializeView

    Me.Show
    IView_ShowDialog = Not This.IsCancelled
End Function


Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub
