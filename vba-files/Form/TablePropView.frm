VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TablePropView 
   Caption         =   "Table Properties"
   ClientHeight    =   8895.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "TablePropView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TablePropView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements iview

Private WithEvents mViewModel As TablePropViewModel

Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdApply_Click()
    mViewModel.ApplyViewModelCommand.Execute
    UpdateControls
End Sub

Private Sub cmdCancel_Click()
    OnCancel
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub cmdActivateListObject_Click()
    mViewModel.ActivateListObjectCommand.Execute
    UpdateControls
End Sub

Private Sub cmdResetValueColumns_Click()
    mViewModel.ColumnProperties.Reset
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set mViewModel = ViewModel
    This.IsCancelled = False
    
    Initalize
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub Initalize()
    UpdateControls
    InitializeLabelPictures
    mViewModel.ColumnProperties.LoadListView Me.lvStarredColumns
    mViewModel.KeyColumns.LoadComboBox Me.cboPreferKeyColumn
End Sub

Private Sub UpdateControls()
    Me.txtTableName.Value = mViewModel.TableName
    Me.txtWorkSheetName.Value = mViewModel.WorkSheetName
    Me.txtWorkBookName.Value = mViewModel.WorkBookName
    
    Me.optLocationLocal = (mViewModel.WorkbookProperty.StorageLocation = LocalStorage)
    Me.optLocationNetwork = (mViewModel.WorkbookProperty.StorageLocation = RemoteStorage)
    Me.optLocationOneDrive = (mViewModel.WorkbookProperty.StorageLocation = OneDriveStorage)
    Me.optLocationSharePoint = (mViewModel.WorkbookProperty.StorageLocation = SharePointStorage)
    
    mViewModel.ActivateListObjectCommand.UpdateCommandButton Me.cmdActivateListObject
    mViewModel.ApplyViewModelCommand.UpdateCommandButton Me.cmdApply
End Sub

Private Sub InitializeLabelPictures()
    Set Me.lblPicDetails.Picture = Application.CommandBars.GetImageMso("TablePropertiesDialog", 32, 32)
    Set Me.lblPicDirection.Picture = Application.CommandBars.GetImageMso("RelationshipsEditRelationships", 32, 32)
    Set Me.lblPicHighlighting.Picture = Application.CommandBars.GetImageMso("HighlightFilters", 32, 32)
    Set Me.lblPicKey.Picture = Application.CommandBars.GetImageMso("AdpDiagramKeys", 32, 32)
    Set Me.lblPicLocation.Picture = Application.CommandBars.GetImageMso("FileFind", 32, 32)
    Set Me.lblPicProfile.Picture = Application.CommandBars.GetImageMso("EnterpriseProjectProfiles", 32, 32)
    Set Me.lblPicProtection.Picture = Application.CommandBars.GetImageMso("SheetProtect", 32, 32)
    Set Me.lblPicTimestamp.Picture = Application.CommandBars.GetImageMso("ViewAllProposals", 32, 32)
End Sub

Private Sub lvStarredColumns_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mViewModel.ColumnProperties.TrySelectByName(Item.Text) Then
        mViewModel.IsDirty = True
        mViewModel.ApplyViewModelCommand.UpdateCommandButton Me.cmdApply
    End If
End Sub

Private Sub cboPreferKeyColumn_Change()
    If mViewModel.KeyColumns.TrySelectByName(Me.cboPreferKeyColumn.Value) Then
        mViewModel.IsDirty = True
        mViewModel.ApplyViewModelCommand.UpdateCommandButton Me.cmdApply
    End If
End Sub

Private Sub chkDeferActivationOnWorksheet_Click()
    mViewModel.ColumnProperties.ActivateOnWorksheet = Me.chkDeferActivationOnWorksheet.Value
End Sub
