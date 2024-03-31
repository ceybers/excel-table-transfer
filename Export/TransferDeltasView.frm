VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferDeltasView 
   Caption         =   "Preview Transfer Changes"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "TransferDeltasView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "TransferDeltasView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@IgnoreModule HungarianNotation
'@Folder "MVVM2.Views"
Option Explicit
Implements IView2

Private Const LBL_HEADING As String = "Table differences compared successfully. The changes can be previewed below before applying them to the Destination table."
Private Const LBL_NO_DELTAS As String = "Table differences compared successfully." & vbCrLf & "No changes were found."

Private Type TState
    ViewModel As TransferDeltasViewModel
    Result As ViewResult
End Type
Private This As TState

Private Sub cmbBack_Click()
    This.Result = vrBack
    Me.Hide
End Sub

Private Sub cmbCancel_Click()
    This.Result = vrCancel
    Me.Hide
End Sub

Private Sub cmbFinish_Click()
    This.Result = vrFinish
    Me.Hide
End Sub

Private Sub cmbNext_Click()
    This.Result = vrNext
    Me.Hide
End Sub

Private Sub cmbStart_Click()
    This.Result = vrStart
    Me.Hide
End Sub

Private Sub cmbShowAll_Click()
    DoShowAll
End Sub

Private Sub lblTargetIcon_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Private Sub lvKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
    DoSelectKey Item.Text
End Sub

Private Sub lvFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    DoSelectField Item.Text
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
    UpdateNoChanges
    
    Me.Show
    
    IView2_ShowDialog = This.Result 'Not This.IsCancelled
End Function

Private Sub InitializeControls()
    Set Me.lblTargetIcon.Picture = frmPictures32.lblCompareChanges.Picture
    
    Me.lblHeading.Caption = LBL_HEADING
    
    TransferDeltasToListView.Initialize Me.lvKeys, tdKeyMember
    TransferDeltasToListView.Initialize Me.lvFields, tdField
    TransferDeltasToListView.Initialize Me.lvDeltas, tdDelta
End Sub

Private Sub UpdateButtons()
    Me.cmbShowAll.Enabled = This.ViewModel.CanShowAll
    Me.cmbNext.Enabled = False ' This is the last step, user will click 'Finish' instead
    Me.cmbFinish.Enabled = This.ViewModel.CanFinish
    
    If Me.cmbFinish.Enabled Then
        Me.cmbFinish.SetFocus
    End If
End Sub

Private Sub UpdateNoChanges()
    If This.ViewModel.CanFinish = False Then
        Me.lblHeading.Caption = LBL_NO_DELTAS
    End If
End Sub

Private Sub UpdateListViewLHS()
    TransferDeltasToListView.Load Me.lvKeys, This.ViewModel, tdKeyMember
    TransferDeltasToListView.Load Me.lvFields, This.ViewModel, tdField
End Sub

Private Sub UpdateListViewRHS()
    TransferDeltasToListView.Load Me.lvDeltas, This.ViewModel, tdDelta
End Sub

Private Sub DoShowAll()
    'msgbox "Are you sure?"
    This.ViewModel.DoShowAll
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub DoSelectKey(ByVal Key As String)
    If Key = TransferDeltasToListView.SELECT_ALL Then
        This.ViewModel.TrySelectKey vbNullString
    Else
        This.ViewModel.TrySelectKey Key
    End If
    
    UpdateListViewRHS
End Sub

Private Sub DoSelectField(ByVal Field As String)
    If Field = TransferDeltasToListView.SELECT_ALL Then
        This.ViewModel.TrySelectField vbNullString
    Else
        This.ViewModel.TrySelectField Field
    End If
    
    UpdateListViewRHS
End Sub
