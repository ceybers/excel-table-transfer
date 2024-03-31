VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TablePickerView 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "TablePickerView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "TablePickerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("MVVM2.Views")
Option Explicit
Implements IView2

Private Const LBL_HEADER_TEXT As String = "Which two tables are you transferring data between?" & vbCrLf & vbCrLf & "Select a table that will be the Source of your data and table that will be the Destination that will be updated."

Private Type TState
    ViewModel As TablePickerViewModel
    Result As ViewResult
End Type
Private This As TState

Private Sub cboBack_Click()
    This.Result = vrBack
    Me.Hide
End Sub

Private Sub cboNext_Click()
    This.Result = vrNext
    Me.Hide
End Sub

Private Sub cboCancel_Click()
    This.Result = vrCancel
    Me.Hide
End Sub

Private Sub cboSelSrc_Click()
    This.ViewModel.DoSelect tdSource
    UpdateListViewLHS ' To update TreeView icons
    UpdateListViewRHS
    UpdateButtons ' Check to Set focus to cboNext
End Sub

Private Sub cboSelDst_Click()
    This.ViewModel.DoSelect tdDestination
    UpdateListViewLHS ' To update TreeView icons
    UpdateListViewRHS
    UpdateButtons ' Check to Set focus to cboNext
End Sub

Private Sub tvTables_NodeClick(ByVal Node As MSComctlLib.Node)
    If This.ViewModel.TryClick(Node.Key) Then
        'UpdateButtons
        ' Even if it returns false we need to disable buttons
    End If
    UpdateButtons
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Function IView2_ShowDialog(ByVal ViewModel As Object) As ViewResult
    Set This.ViewModel = ViewModel
    
    InitializeControls
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    
    Me.Show
    
    IView2_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    Me.lblHeaderText.Caption = LBL_HEADER_TEXT
    TablePickerToTreeView.Initialize Me.tvTables
End Sub

Private Sub UpdateListViewLHS()
    TablePickerToTreeView.Load Me.tvTables, This.ViewModel
End Sub

Private Sub UpdateListViewRHS()
    TablePickerToFrame.UpdateControls This.ViewModel, Me.frmSrc, tdSource
    TablePickerToFrame.UpdateControls This.ViewModel, Me.frmDst, tdDestination
End Sub

Private Sub UpdateButtons()
    Me.cboBack.Enabled = False
    Me.cboNext.Enabled = This.ViewModel.CanNext
    Me.cboCancel.Enabled = True
    
    Me.cboSelSrc.Enabled = This.ViewModel.CanSelect
    Me.cboSelDst.Enabled = This.ViewModel.CanSelect
    
    If This.ViewModel.CanNext Then
        Me.cboNext.SetFocus
    End If
End Sub
