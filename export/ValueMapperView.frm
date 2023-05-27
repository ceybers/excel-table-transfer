VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValueMapperView 
   Caption         =   "Select Columns to Map"
   ClientHeight    =   8205.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   OleObjectBlob   =   "ValueMapperView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValueMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.ValueMapper.View"
Option Explicit
Implements IView
Implements ICancellable
 
Private Type TView
    Context As IAppContext
    IsCancelled As Boolean
    ViewModel As ValueMapperViewModel
End Type
Private This As TView

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As ValueMapperViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As ValueMapperViewModel)
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

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As ValueMapperViewModel) As IView
    Dim Result As ValueMapperView
    Set Result = New ValueMapperView
    
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
    
    ViewModel.ExtractMappedData
    
    IView_ShowDialog = Not This.IsCancelled
End Function


Private Sub InitializeControls()
    '@Ignore ArgumentWithIncompatibleObjectType
     SrcValueColumnsToListViewConv.InitializeListView Me.lvSrcValueColumns
    '@Ignore ArgumentWithIncompatibleObjectType
     DstValueColumnsToListViewConv.InitializeListView Me.lvDstValueColumns
End Sub

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "SrcValueColumnVM.SelectedKey", Me.txtSrcSelected, "Value", OneWayBinding
        .BindPropertyPath ViewModel, "DstValueColumnVM.SelectedKey", Me.txtDstSelected, "Value", OneWayBinding
        
        .BindPropertyPath ViewModel, "EnableStarredOnly", Me.chkEnableStarredOnly, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "EnableOption2", Me.chkEnableOption2, "Value", TwoWayBinding
        
        
        .BindPropertyPath ViewModel, "SrcValueColumnVM.TableName", Me.txtSrcTableProps, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "SrcValueColumnVM.Item", Me.lvSrcValueColumns, "ListItems", TwoWayBinding, SrcValueColumnsToListViewConv
        
        .BindPropertyPath ViewModel, "DstValueColumnVM.TableName", Me.txtDstTableProps, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "DstValueColumnVM.Item", Me.lvDstValueColumns, "ListItems", TwoWayBinding, DstValueColumnsToListViewConv
    End With
End Sub

Private Sub BindCommands()
    BindCommand OKViewCommand.Create(Context, Me, ViewModel), Me.cmdOK
    BindCommand CancelViewCommand.Create(Context, Me, ViewModel), Me.cmdCancel
    
    BindCommand MapValueColumnsCommand.Create(Context, Me, ViewModel), Me.cmbMapValueColumns
    BindCommand UnmapValueColumnsCommand.Create(Context, Me, ViewModel), Me.cmbUnmapValueColumns
    BindCommand UnmapAllColumnsCommand.Create(Context, Me, ViewModel), Me.cmbUnmapAll
    BindCommand AutoMapColumnsCommand.Create(Context, Me, ViewModel), Me.cmbAutoMap
    
    BindCommand ShowTablePropsCommand.Create(Context, Me, ViewModel.SrcValueColumnVM), Me.cmbSrcTableProps
    BindCommand ShowTablePropsCommand.Create(Context, Me, ViewModel.DstValueColumnVM), Me.cmbDstTableProps
End Sub

Private Sub BindCommand(ByVal Command As ICommand, ByVal Control As Object)
    This.Context.CommandManager.BindCommand Context, ViewModel, Command, Control
End Sub

Private Sub InitializeLabelPictures()
    ApplyImageMSOtoLabel Me.lblPicHeader, "TablePropertiesDialog"
    ApplyImageMSOtoLabel Me.lblPicOptions, "TablePropertiesDialog"
    ApplyImageMSOtoLabel Me.lblPicSrcTable, "TablePropertiesDialog"
    ApplyImageMSOtoLabel Me.lblPicDstTable, "TablePropertiesDialog"
End Sub
