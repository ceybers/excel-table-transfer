VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyMapper2View 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "KeyMapper2View.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "KeyMapper2View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("MVVM2.Views")
Option Explicit
Implements IView2

Private Const LBL_HEADER_TEXT As String = "Which two columns contain the primary key used to join your tables?" & vbCrLf & vbCrLf & _
    "Select a Key column in both the Source and the Destination table."

Private Type TState
    ViewModel As KeyMapper2ViewModel
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

Private Sub lvSrcKeys_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    UpdateButtons
End Sub

Private Sub lvDstKeys_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    UpdateButtons
End Sub

Private Sub cmbDstQuality_DropButtonClick()
    Me.cmbDstQuality.Enabled = False
    Me.cmbDstQuality.Enabled = True
    This.ViewModel.ShowQuality TransferDirection.tdDestination

End Sub

Private Sub cmbMatchQuality_DropButtonClick()
    Me.cmbMatchQuality.Enabled = False
    Me.cmbMatchQuality.Enabled = True
    This.ViewModel.ShowMatchDetails
End Sub

Private Sub cmbSrcQuality_DropButtonClick()
    Me.cmbSrcQuality.Enabled = False
    Me.cmbSrcQuality.Enabled = True
    This.ViewModel.ShowQuality TransferDirection.tdSource
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
    UpdateResults
    UpdateButtons
    
    Me.Show
    
    IView2_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    Me.lblHeaderText.Caption = LBL_HEADER_TEXT
    
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
