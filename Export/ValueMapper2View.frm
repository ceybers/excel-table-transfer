VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValueMapper2View 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "ValueMapper2View.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ValueMapper2View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("MVVM2.Views")
Option Explicit
Implements IView2

Private Const LBL_HEADER_TEXT As String = "Which columns contain the data you want to transfer?" & vbCrLf & vbCrLf & _
    "Select pairs of columns from the Source and Destination tables and add them to the Mapped columns."

Private Type TState
    ViewModel As ValueMapper2ViewModel
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

Private Function IView2_ShowDialog(ByVal ViewModel As Object) As ViewResult
    Set This.ViewModel = ViewModel
    
    InitializeControls
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    Me.cboMapAll.SetFocus
    
    Me.Show
    
    IView2_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    Me.lblHeaderText.Caption = LBL_HEADER_TEXT
    
    ValueColumnsToListView.Initialize Me.lvSrcValues
    ValueColumnsToListView.Initialize Me.lvDstValues
    ColumnPairs2ToListView.Initialize Me.lvMappedValues
End Sub

Private Sub UpdateListViewLHS()
    ValueColumnsToListView.Load Me.lvSrcValues, This.ViewModel.Source
End Sub

Private Sub UpdateListViewRHS()
    ValueColumnsToListView.Load Me.lvDstValues, This.ViewModel.Destination
End Sub

Private Sub UpdateMappingControl()
    ColumnPairs2ToListView.Load Me.lvMappedValues, This.ViewModel.Mapped
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


