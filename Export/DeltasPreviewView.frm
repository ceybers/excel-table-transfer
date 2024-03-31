VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeltasPreviewView 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "DeltasPreviewView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "DeltasPreviewView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule HungarianNotation
'@Folder "MVVM.Views"
Option Explicit
Implements IView

Private Type TState
    ViewModel As DeltasPreviewViewModel
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

Private Sub lblHeaderIcon_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As ViewResult
    Set This.ViewModel = ViewModel
    
    InitializeControls
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    UpdateNoChanges
    
    Me.Show
    
    IView_ShowDialog = This.Result 'Not This.IsCancelled
End Function

Private Sub InitializeControls()
    Me.lblHeaderText.Caption = HDR_TXT_DELTAS_PREVIEW
    
    TransferDeltasToListView.Initialize Me.lvKeys, ttKeyMember
    TransferDeltasToListView.Initialize Me.lvFields, ttField
    TransferDeltasToListView.Initialize Me.lvDeltas, ttDelta
End Sub

Private Sub UpdateButtons()
    Me.cmbShowAll.Enabled = This.ViewModel.CanShowAll
    Me.cmbNext.Enabled = This.ViewModel.CanFinish
    
    If Me.cmbNext.Enabled Then
        Me.cmbNext.SetFocus
    End If
End Sub

Private Sub UpdateNoChanges()
    If This.ViewModel.CanFinish = False Then
        Me.lblHeaderText.Caption = DELTAS_PREVIEW_NO_RESULTS
    End If
End Sub

Private Sub UpdateListViewLHS()
    TransferDeltasToListView.Load Me.lvKeys, This.ViewModel, ttKeyMember
    TransferDeltasToListView.Load Me.lvFields, This.ViewModel, ttField
End Sub

Private Sub UpdateListViewRHS()
    TransferDeltasToListView.Load Me.lvDeltas, This.ViewModel, ttDelta
End Sub

Private Sub DoShowAll()
    'msgbox "Are you sure?"
    This.ViewModel.DoShowAll
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub DoSelectKey(ByVal Key As String)
    If Key = SELECT_ALL Then
        This.ViewModel.TrySelectKey vbNullString
    Else
        This.ViewModel.TrySelectKey Key
    End If
    
    UpdateListViewRHS
End Sub

Private Sub DoSelectField(ByVal Field As String)
    If Field = SELECT_ALL Then
        This.ViewModel.TrySelectField vbNullString
    Else
        This.ViewModel.TrySelectField Field
    End If
    
    UpdateListViewRHS
End Sub
