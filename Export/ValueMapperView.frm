VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValueMapperView 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "ValueMapperView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ValueMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'@Folder "MVVM.Views"
Option Explicit
Implements IView

Private Type TState
    ViewModel As ValueMapperViewModel
    Result As TtViewResult
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

Private Sub lvSrcValues_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.Source.TrySelect Item.Key
    'This.ViewModel.TryEvaluateMatch
    UpdateButtons
    UpdateMappingControl
End Sub

Private Sub lvDstValues_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.Destination.TrySelect Item.Key
    'This.ViewModel.TryEvaluateMatch
    UpdateButtons
    UpdateMappingControl
End Sub

Private Sub cboMapAdd_Click()
    This.ViewModel.DoMapAdd
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateMappingControl
    UpdateButtons
End Sub

Private Sub cboMapRemove_Click()
    This.ViewModel.DoMapRemove
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateMappingControl
    UpdateButtons
End Sub

Private Sub cboMapAll_Click()
    This.ViewModel.DoAutoMap
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateMappingControl
    UpdateButtons
    
    If Me.cboNext.Enabled Then Me.cboNext.SetFocus
End Sub

Private Sub cboMapReset_Click()
    This.ViewModel.DoRemoveAll
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateMappingControl
    UpdateButtons
    
    Me.cboMapAll.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Sub lblHeaderIcon_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As TtViewResult
    Set This.ViewModel = ViewModel
    
    InitializeControls
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    Me.cboMapAll.SetFocus
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    Me.lblHeaderText.Caption = HDR_TXT_VALUE_MAPPER
    
    ValueColumnsToListView.Initialize Me.lvSrcValues
    ValueColumnsToListView.Initialize Me.lvDstValues
    ColumnPairsToListView.Initialize Me.lvMappedValues
End Sub

Private Sub UpdateListViewLHS()
    ValueColumnsToListView.Load Me.lvSrcValues, This.ViewModel.Source
End Sub

Private Sub UpdateListViewRHS()
    ValueColumnsToListView.Load Me.lvDstValues, This.ViewModel.Destination
End Sub

Private Sub UpdateMappingControl()
    ColumnPairsToListView.Load Me.lvMappedValues, This.ViewModel.Mapped
End Sub

Private Sub UpdateButtons()
    Me.cboBack.Enabled = True
    Me.cboNext.Enabled = This.ViewModel.CanNext
    Me.cboCancel.Enabled = True
    
    Me.cboMapAdd.Enabled = This.ViewModel.CanMapAdd
    Me.cboMapRemove.Enabled = This.ViewModel.CanMapRemove
    Me.cboMapAll.Enabled = True
    Me.cboMapReset.Enabled = This.ViewModel.CanRemoveAll
End Sub

