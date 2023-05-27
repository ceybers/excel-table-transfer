VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableMapperView 
   Caption         =   "Select Source and Destination Tables"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10710
   OleObjectBlob   =   "TableMapperView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TableMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.TableMapper.View"
Option Explicit
Implements IView
Implements ICancellable
 
Private Type TView
    Context As IAppContext
    IsCancelled As Boolean
    ViewModel As TableMapperViewModel
End Type
Private This As TView

Private Sub cmdDEBUG_Click()
    Debug.Print "..."
    Debug.Print ".."
    Debug.Print "."
End Sub

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As TableMapperViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As TableMapperViewModel)
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

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As TableMapperViewModel) As IView
    Dim Result As TableMapperView
    Set Result = New TableMapperView
    
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
     PickableTablesToListViewConv.InitializeListView Me.lvSrcTables
    '@Ignore ArgumentWithIncompatibleObjectType
     PickableTablesToListViewConv.InitializeListView Me.lvDstTables
End Sub

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "EnableProtected", Me.chkEnableProtected, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "EnableReadonly", Me.chkEnableReadOnly, "Value", TwoWayBinding
        
        .BindPropertyPath ViewModel, "SrcTableVM.Item", Me.lvSrcTables, "ListItems", TwoWayBinding, PickableTablesToListViewConv
        .BindPropertyPath ViewModel, "DstTableVM.Item", Me.lvDstTables, "ListItems", TwoWayBinding, PickableTablesToListViewConv
    End With
End Sub

Private Sub BindCommands()
    Dim OKView As ICommand
    Set OKView = OKViewCommand.Create(Context, Me, ViewModel)
    
    Dim CancelView As ICommand
    Set CancelView = CancelViewCommand.Create(Context, Me, ViewModel)
    
    'Dim SrcDblClickListView As ICommand
    'Set SrcDblClickListView = DblClickListViewCommand.Create(Context, Me, ViewModel.SrcTableVM)
    
    With This.Context.CommandManager
        .BindCommand Context, ViewModel, OKView, Me.cmdOK
        .BindCommand Context, ViewModel, CancelView, Me.cmdCancel
        '.BindCommand Context, ViewModel.SrcTableVM, SrcDblClickListView, Me.lvSrcTables
    End With
End Sub

Private Sub InitializeLabelPictures()
    ApplyImageMSOtoLabel Me.lblPicHeader, "TablePropertiesDialog"
    ApplyImageMSOtoLabel Me.lblPicOptions, "TablePropertiesDialog"
End Sub
