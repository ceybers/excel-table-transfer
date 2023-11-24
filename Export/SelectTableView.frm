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
    IsCancelled As Boolean
End Type

Private this As TView

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
    vm.LoadTreeview Me.tvTables
End Sub

Private Sub vm_ItemSelected()
    If vm.SelectedTable Is Nothing Then
        Me.cmbOK.Enabled = False
    Else
        Me.cmbOK.Enabled = True
    End If
End Sub

Private Sub txtSearch_Change()
    vm.SearchCriteria = (txtSearch & "*")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As IViewModel) As Boolean
    Set vm = ViewModel
    
    InitializeView
    
    If vm.AutoSelected = True Then
        IView_ShowDialog = True
        Exit Function
    End If
    
    If vm.CanSelect Then
        Me.Show
    End If
    
    IView_ShowDialog = Not this.IsCancelled
End Function

Private Sub InitializeView()
    this.IsCancelled = False
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
        If nd.key = vm.SelectedTable.Range.Address(external:=True) Then
            nd.Selected = True
            nd.EnsureVisible

            Me.cmbOK.Enabled = True
            Me.cmbOK.SetFocus
            Exit Sub
        End If
    Next nd
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub
