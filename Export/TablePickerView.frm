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




'@Folder "MVVM.Views"
Option Explicit
Implements IView

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
    This.ViewModel.PickSelectedTable ttSource
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub cboSelDst_Click()
    This.ViewModel.PickSelectedTable ttDestination
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub tvTables_NodeClick(ByVal Node As MSComctlLib.Node)
    This.ViewModel.TrySelect Node.Key
    UpdateButtons
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Sub lblHeaderIcon_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As ViewResult
    Set This.ViewModel = ViewModel
    
    InitializeControls
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    Me.lblHeaderText.Caption = HDR_TXT_TABLE_PICKER
    TablePickerToTreeView.Initialize Me.tvTables
End Sub

Private Sub UpdateListViewLHS()
    TablePickerToTreeView.Load Me.tvTables, This.ViewModel
End Sub

Private Sub UpdateListViewRHS()
    TablePickerToFrame.Load This.ViewModel, Me.frmSrc, ttSource
    TablePickerToFrame.Load This.ViewModel, Me.frmDst, ttDestination
End Sub

Private Sub UpdateButtons()
    Me.cboBack.Enabled = False
    Me.cboNext.Enabled = This.ViewModel.CanNext
    Me.cboCancel.Enabled = True
    
    Me.cboSelSrc.Enabled = This.ViewModel.CanPickSelected
    Me.cboSelDst.Enabled = This.ViewModel.CanPickSelected
    
    If This.ViewModel.CanNext Then
        Me.cboNext.SetFocus
    End If
End Sub

