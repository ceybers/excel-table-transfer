VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectTableView 
   Caption         =   "Select Table"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "SelectTableView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "SelectTableView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@IgnoreModule HungarianNotation
'@Folder "MVVM.Views"
Option Explicit
Implements IView2

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As SelectTableViewModel
Attribute vm.VB_VarHelpID = -1
Private Const ICON_SIZE As Long = 16

Private Type TView
    Result As ViewResult
End Type
Private This As TView

Private Sub cmbClearSearch_Click()
    Me.txtSearch = vbNullString
    vm.SearchCriteria = vbNullString
    Me.txtSearch.SetFocus
End Sub

Private Sub cmbOK_Click()
    This.Result = vrNext
    Me.Hide
End Sub

Private Sub tvTables_DblClick()
    If Not vm.SelectedTable Is Nothing Then
        This.Result = vrNext
        Me.Hide
    End If
End Sub

Private Sub tvTables_NodeClick(ByVal Node As MSComctlLib.Node)
    vm.TrySelect Node.Key
End Sub

Private Sub vm_CollectionChanged()
    vm.LoadTreeview Me.tvTables
End Sub

Private Sub vm_ItemSelected()
    Me.cmbOK.Enabled = Not (vm.SelectedTable Is Nothing)
End Sub

Private Sub txtSearch_Change()
    vm.SearchCriteria = (txtSearch & "*")
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Function IView2_ShowDialog(ByVal ViewModel As Object) As ViewResult
    Set vm = ViewModel
    
    InitializeView
    
    If vm.AutoSelected = True Then
        IView2_ShowDialog = vrNext
        Exit Function
    End If
    
    If vm.CanSelect Then
        Me.Show
    End If
    
    IView2_ShowDialog = This.Result
End Function

Private Sub InitializeView()
    Me.txtSearch = vbNullString
    Me.cmbOK.Enabled = False
    
    vm.LoadTreeview Me.tvTables
    
    Set Me.cmbClearSearch.Picture = GetMSOImageList(ICON_SIZE).ListImages.Item("delete").Picture
    Me.txtSearch.SetFocus
    
    UpdateListViewWithSelectedTable
End Sub

Private Sub UpdateListViewWithSelectedTable()
    If vm.SelectedTable Is Nothing Then
        Exit Sub
    End If
    
    Dim nd As Node
    For Each nd In Me.tvTables.Nodes
        If nd.Key = vm.SelectedTable.Range.Address(External:=True) Then
            nd.Selected = True
            nd.EnsureVisible

            Me.cmbOK.Enabled = True
            Me.cmbOK.SetFocus
            Exit Sub
        End If
    Next nd
End Sub

