VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyMapperView 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "KeyMapperView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "KeyMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'@Folder "MVVM.Views"
Option Explicit
Implements IView

Private Type TState
    ViewModel As KeyMapperViewModel
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

Private Sub lvSrcKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.Source.TrySelect Item.Key
    This.ViewModel.TryEvaluateMatch
    
    UpdateListViewLHS
    UpdateResults
End Sub

Private Sub lvDstKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.Destination.TrySelect Item.Key
    This.ViewModel.TryEvaluateMatch
    
    UpdateListViewRHS
    UpdateResults
End Sub

Private Sub lvSrcKeys_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    UpdateButtons
End Sub

Private Sub lvDstKeys_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    UpdateButtons
End Sub

Private Sub cmbSrcQuality_DropButtonClick()
    Me.cboCancel.SetFocus
    This.ViewModel.ShowQuality TransferDirection.ttSource
End Sub

Private Sub cmbDstQuality_DropButtonClick()
    Me.cboCancel.SetFocus
    This.ViewModel.ShowQuality TransferDirection.ttDestination
End Sub

Private Sub cmbMatchQuality_DropButtonClick()
    Me.cboCancel.SetFocus
    This.ViewModel.ShowMatchDetails
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
    UpdateResults
    UpdateButtons
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    Me.lblHeaderText.Caption = HDR_TXT_KEY_MAPPER
    
    KeyColumnsToListView.Initialize Me.lvSrcKeys
    KeyColumnsToListView.Initialize Me.lvDstKeys
End Sub

Private Sub UpdateListViewLHS()
    KeyColumnsToListView.Load Me.lvSrcKeys, This.ViewModel.Source
    KeyColumnsToComboBox.Load Me.cmbSrcQuality, This.ViewModel.Source
End Sub

Private Sub UpdateListViewRHS()
    KeyColumnsToListView.Load Me.lvDstKeys, This.ViewModel.Destination
    KeyColumnsToComboBox.Load Me.cmbDstQuality, This.ViewModel.Destination
End Sub

Private Sub UpdateResults()
    MatchQualityToComboBox.Load Me.cmbMatchQuality, This.ViewModel.MatchQuality
End Sub

Private Sub UpdateButtons()
    Me.cboBack.Enabled = True
    Me.cboNext.Enabled = This.ViewModel.CanNext
    Me.cboCancel.Enabled = True
    
    Me.cmbSrcQuality.Enabled = Not This.ViewModel.Source.Selected Is Nothing
    Me.cmbDstQuality.Enabled = Not This.ViewModel.Destination.Selected Is Nothing
    Me.cmbMatchQuality.Enabled = Not This.ViewModel.MatchQuality Is Nothing
    
    If This.ViewModel.CanNext Then
        Me.cboNext.SetFocus
    End If
End Sub

