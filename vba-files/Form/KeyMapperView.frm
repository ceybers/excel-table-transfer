VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyMapperView 
   Caption         =   "KeyMapperView"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10650
   OleObjectBlob   =   "KeyMapperView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KeyMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.KeyMapper.View"
Option Explicit
Implements IView
Implements ICancellable
 
Private Type TView
    Context As IAppContext
    IsCancelled As Boolean
    ViewModel As KeyMapperViewModel
End Type
Private This As TView

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As KeyMapperViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As KeyMapperViewModel)
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

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As KeyMapperViewModel) As IView
    Dim Result As KeyMapperView
    Set Result = New KeyMapperView
    
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
     KeyColumnsToListViewConv.InitializeListView Me.lvSrcKeyColumns
    '@Ignore ArgumentWithIncompatibleObjectType
     KeyColumnsToListViewConv.InitializeListView Me.lvDstKeyColumns
End Sub

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "SrcKeyColumnVM.SelectedAsText", Me.TextBox1, "Value", OneWayBinding
        
        .BindPropertyPath ViewModel, "EnableNontext", Me.chkEnableNontext, "Value", TwoWayBinding
        .BindPropertyPath ViewModel, "EnableNonunique", Me.chkEnableNonunique, "Value", TwoWayBinding
        
        .BindPropertyPath ViewModel, "SrcKeyColumnVM.Item", Me.lvSrcKeyColumns, "ListItems", TwoWayBinding, KeyColumnsToListViewConv
        .BindPropertyPath ViewModel, "DstKeyColumnVM.Item", Me.lvDstKeyColumns, "ListItems", TwoWayBinding, KeyColumnsToListViewConv
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
    ApplyImageMSOtoLabel Me.lblPicHeader, "TablePropertiesDialog"
    ApplyImageMSOtoLabel Me.lblPicOptions, "TablePropertiesDialog"
End Sub

