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
'@Folder "MVVM.TableProps.View"
Option Explicit
Implements IView
Implements ICancellable
 
Private Type TView
    Context As IAppContext
    IsCancelled As Boolean
    ViewModel As TablePropViewModel
End Type
Private This As TView

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As TablePropViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As TablePropViewModel)
    Set This.ViewModel = vNewValue
End Property

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set This.Context = vNewValue
End Property

' TODO Hacky shim fix
Private Sub mpgTabs_Change()
    If Me.mpgTabs.Value = 1 Then ' second tab has the LV
        This.ViewModel.TableStarColumnsVM.Repaint
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
 
Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Show()
    IView_ShowDialog
End Sub
 
Private Sub IView_Hide()
    Me.Hide
End Sub

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As TablePropViewModel) As IView
    Dim Result As TablePropView
    Set Result = New TablePropView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView_ShowDialog() As Boolean
    InitializeLabelPictures
    InitializeControls
    BindControls
    BindCommands
    
    Me.Show vbModal
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitializeControls()
    '@Ignore ArgumentWithIncompatibleObjectType
    ColumnPropToListViewConverter.InitializeListView This.Context, Me.lvStarredColumns
    
    ' Activate first tab of multipage
    Me.mpgTabs.Value = 0
End Sub

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "TableDetailsVM.TableName", Me.txtTableName, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "TableDetailsVM.WorkSheetName", Me.txtWorkSheetName, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "TableDetailsVM.WorkBookName", Me.txtWorkBookName, "Value", OneTimeBinding
        
        .BindPropertyPath ViewModel, "TableLocationVM.IsLocalStorage", Me.optLocationLocal, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "TableLocationVM.IsNetworkStorage", Me.optLocationNetwork, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "TableLocationVM.IsOneDriveStorage", Me.optLocationOneDrive, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "TableLocationVM.IsSharePointStorage", Me.optLocationSharePoint, "Value", OneTimeBinding
        
        .BindPropertyPath ViewModel, "TableDirectionVM.IsNeither", Me.optDirectionNeither, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableDirectionVM.IsSource", Me.optDirectionSource, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableDirectionVM.IsDestination", Me.optDirectionDestination, "Value", TwoWayBinding
        
        .BindPropertyPath ViewModel, "TableDirectionVM.ConditionDirections", Me.cboPreferDirectionLocation, "List", OneTimeBinding ' Must be first for this control
        .BindPropertyPath ViewModel, "TableDirectionVM.IsOnCondition", Me.chkPreferDirectionCondition, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableDirectionVM.IsOnCondition", Me.cboPreferDirectionLocation, "Enabled", TwoWayBinding
        .BindPropertyPath ViewModel, "TableDirectionVM.SelectedConditionDirection", Me.cboPreferDirectionLocation, "Value", TwoWayBinding
        
        .BindPropertyPath ViewModel, "TablePreferKeyVM.Columns", Me.cboPreferKeyColumn, "List", OneTimeBinding
        .BindPropertyPath ViewModel, "TablePreferKeyVM.SelectedColumn", Me.cboPreferKeyColumn, "Value", TwoWayBinding
        
        .BindPropertyPath ViewModel, "TableStarColumnsVM.Columns", Me.lvStarredColumns, "ListItems", TwoWayBinding, ColumnPropToListViewConverter
        
        .BindPropertyPath ViewModel, "TableTimestampVM.IsEnabled", Me.chkEnableTimestamp, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableTimestampVM.Address", Me.txtTimestampCell, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableTimestampVM.IsEnabled", Me.txtTimestampCell, "Enabled", OneWayBinding
        
        .BindPropertyPath ViewModel, "TableProtectionVM.IsNoChange", Me.optProtectionNoChange, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableProtectionVM.IsTemporary", Me.optProtectionTemporarily, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableProtectionVM.IsPermanent", Me.optProtectionPermanently, "Value", TwoWayBinding
        
        .BindPropertyPath ViewModel, "TableProtectionVM.IsTableProtected", Me.chkProtectionProtect, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableProtectionVM.IsPasswordProtected", Me.chkProtectionPassword, "Value", OneWayBinding
        
        .BindPropertyPath ViewModel, "TableProfileVM.HasProfile", Me.cmdProfileRemove, "Enabled", OneTimeBinding
    End With
End Sub

Private Sub BindCommands()
    BindCommand OKViewCommand.Create(Context, Me, ViewModel), Me.cmdOK
    BindCommand CancelViewCommand.Create(Context, Me, ViewModel), Me.cmdCancel
    BindCommand ResetStarColumnsCommand.Create(Context, Me, ViewModel.TableStarColumnsVM), Me.cmdResetValueColumns
    BindCommand ActivateProfileCommand.Create(Context, Me, ViewModel.TableProfileVM), Me.cmdProfileActivate
    BindCommand RemoveProfileCommand.Create(Context, Me, ViewModel.TableProfileVM), Me.cmdProfileRemove
    BindCommand RemoveHighlightingCommand.Create(Context, Me, ViewModel.TableHighlightingVM), Me.cmdHighlightingRemove
End Sub

Private Sub BindCommand(ByVal Command As ICommand, ByVal Control As Object)
    This.Context.CommandManager.BindCommand Context, ViewModel, Command, Control
End Sub

Private Sub InitializeLabelPictures()
    ApplyImageMSOtoLabel Me.lblPicDetails, "TablePropertiesDialog"
    ApplyImageMSOtoLabel Me.lblPicDirection, "RelationshipsEditRelationships"
    ApplyImageMSOtoLabel Me.lblPicHighlighting, "HighlightFilters"
    ApplyImageMSOtoLabel Me.lblPicKey, "AdpDiagramKeys"
    ApplyImageMSOtoLabel Me.lblPicLocation, "FileFind"
    ApplyImageMSOtoLabel Me.lblPicProfile, "EnterpriseProjectProfiles"
    ApplyImageMSOtoLabel Me.lblPicProtection, "SheetProtect"
    ApplyImageMSOtoLabel Me.lblPicTimestamp, "ViewAllProposals"
End Sub
