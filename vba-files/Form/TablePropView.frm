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
    'CountryToListViewConverter.InitializeListView Me.ListView1
    '@Ignore ArgumentWithIncompatibleObjectType
    'CitytoListViewConverter.InitializeListView Me.ListView2
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
        
        .BindPropertyPath ViewModel, "TableDirectionVM.IsOnCondition", Me.chkPreferDirectionCondition, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "TableDirectionVM.IsOnCondition", Me.cboPreferDirectionLocation, "Enabled", TwoWayBinding
        .BindPropertyPath ViewModel, "TableDirectionVM.ConditionDirections", Me.cboPreferDirectionLocation, "List", OneWayToSource
        .BindPropertyPath ViewModel, "TableDirectionVM.SelectedConditionDirection", Me.cboPreferDirectionLocation, "Value", TwoWayBinding
        
        '.BindPropertyPath ViewModel, "Countries", Me.ListView1, "ListItems", OneWayToSource, CountryToListViewConverter
        '.BindPropertyPath ViewModel, "Country", Me.ListView1, "SelectedItem"
        '.BindPropertyPath ViewModel, "Country", Me.TextBox1, "Value"
        
        '.BindPropertyPath ViewModel, "CityViewModel.Cities", Me.ListView2, "ListItems", OneWayToSource, CitytoListViewConverter
        '.BindPropertyPath ViewModel, "CityViewModel.SelectedCityKey", Me.ListView2, "SelectedItem", TwoWayBinding, CitytoListViewConverter
        '.BindPropertyPath ViewModel, "CityViewModel.Cities", Me.ComboBox1, "List", OneWayToSource, CityToComboBoxConverter
        '.BindPropertyPath ViewModel, "CityViewModel.SelectedCityKey", Me.ComboBox1, "Value", TwoWayBinding, CityToComboBoxConverter
        
        '.BindPropertyPath ViewModel, "CityViewModel.SelectedCityKey", Me.Label2, "Caption"
    End With
End Sub

Private Sub BindCommands()
    Dim OKView As ICommand
    Set OKView = OKViewCommand.Create(Context, Me, ViewModel)
    
    Dim CancelView As ICommand
    Set CancelView = CancelViewCommand.Create(Context, Me, ViewModel)
    
    With This.Context.CommandManager
        .BindCommand Context, ViewModel, OKView, Me.cmdOK
        .BindCommand Context, ViewModel, CancelView, Me.cmdCancel
    End With
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
