VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferDeltasView 
   Caption         =   "Preview Transfer Changes"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "TransferDeltasView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransferDeltasView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.Views"
Option Explicit
Implements IView2

Private Const LBL_HEADING As String = "Preview the differences between the tables and the proposed changes before applying them to the Destiantion table."

Private Type TState
    ViewModel As TransferDeltasViewModel
    IsCancelled As Boolean
End Type

Private This As TState

' ---
Private Sub cmbOK_Click()
    Me.Hide
End Sub

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub cmbShowAll_Click()
    TryShowAll
End Sub

Private Sub lvFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Key = TransferDeltasToListView.SELECT_ALL Then
        This.ViewModel.TrySelectField vbNullString
    Else
        This.ViewModel.TrySelectField Item.Text
    End If
    UpdateListViewRHS
End Sub

Private Sub lvKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Key = TransferDeltasToListView.SELECT_ALL Then
        This.ViewModel.TrySelectKey vbNullString
    Else
        This.ViewModel.TrySelectKey Item.Text
    End If
    UpdateListViewRHS
End Sub

'Private Sub tvStates_NodeClick(ByVal Node As MSComctlLib.Node)
'    This.ViewModel.TrySelect Node.Key
'    UpdateListViewRHS
'    UpdateButtons
'End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub UserForm_Activate()
    Set Me.lblTargetIcon.Picture = frmPictures.lblTarget.Picture
End Sub

Private Function IView2_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set This.ViewModel = ViewModel
    This.IsCancelled = False
    
    Me.lblHeading.caption = LBL_HEADING
    
    InitializeControls
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    
    Me.Show
    
    IView2_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitializeControls()
    TransferDeltasToListView.Initialize Me.lvKeys, tdKeyMember
    TransferDeltasToListView.Initialize Me.lvFields, tdField
    TransferDeltasToListView.Initialize Me.lvDeltas, tdDelta
End Sub

Private Sub UpdateButtons()
    Me.cmbShowAll.Enabled = This.ViewModel.CanShowAll
    'Me.cmbSave.Enabled = This.ViewModel.CanSave
    'Me.cmbApply.Enabled = This.ViewModel.CanApply
    'Me.cmbExport.Enabled = This.ViewModel.CanExport
End Sub

Private Sub UpdateListViewLHS()
    TransferDeltasToListView.Load Me.lvKeys, This.ViewModel, tdKeyMember
    TransferDeltasToListView.Load Me.lvFields, This.ViewModel, tdField
End Sub

Private Sub UpdateListViewRHS()
    TransferDeltasToListView.Load Me.lvDeltas, This.ViewModel, tdDelta
End Sub

Private Sub TrySave()
    'This.ViewModel.Save
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryShowAll()
    'msgbox "Are you sure?"
    This.ViewModel.DoShowAll
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub


