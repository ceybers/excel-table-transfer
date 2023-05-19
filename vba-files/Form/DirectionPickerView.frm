VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DirectionPickerView 
   Caption         =   "Choose Direction"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "DirectionPickerView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DirectionPickerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "MVVM.DirectionPicker.View"
Option Explicit
Implements IView
Implements ICancellable
 
Private Type TView
    Context As IAppContext
    IsCancelled As Boolean
    ViewModel As DirectionPickerViewModel
End Type
Private This As TView

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As DirectionPickerViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As DirectionPickerViewModel)
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

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As DirectionPickerViewModel) As IView
    Dim Result As DirectionPickerView
    Set Result = New DirectionPickerView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView_ShowDialog() As Boolean
    InitializeLabelPictures
    BindControls
    BindCommands
    
    Me.Show vbModal
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "TableDetailsVM.TableName", Me.txtTableName, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "TableDetailsVM.WorkSheetName", Me.txtWorkSheetName, "Value", OneTimeBinding
        .BindPropertyPath ViewModel, "TableDetailsVM.WorkBookName", Me.txtWorkBookName, "Value", OneTimeBinding
    
        .BindPropertyPath ViewModel, "TableLocationVM.Locations", Me.cboStorageLocation, "List", OneTimeBinding
        .BindPropertyPath ViewModel, "TableLocationVM.SelectedLocation", Me.cboStorageLocation, "Value", OneTimeBinding
    End With
End Sub

Private Sub BindCommands()
    'Dim OKView As ICommand
    'Set OKView = OKViewCommand.Create(Context, Me, ViewModel)
    
    Dim CancelView As ICommand
    Set CancelView = CancelViewCommand.Create(Context, Me, ViewModel)
    
    Dim DirectionSource As ICommand
    Set DirectionSource = DirectionSourceCommand.Create(Context, Me, ViewModel)
    
    Dim DirectionDestination As ICommand
    Set DirectionDestination = DirectionDestinationCommand.Create(Context, Me, ViewModel)
    
    With This.Context.CommandManager
        '.BindCommand Context, ViewModel, OKView, Me.cmdOK
        .BindCommand Context, ViewModel, CancelView, Me.cmdCancel
        .BindCommand Context, ViewModel, DirectionSource, Me.cmbDirectionSource
        .BindCommand Context, ViewModel, DirectionDestination, Me.cmbDirectionDestination
    End With
End Sub

Private Sub InitializeLabelPictures()
    Set Me.lblPicDetails.Picture = Application.CommandBars.GetImageMso("TablePropertiesDialog", 32, 32)
    Set Me.cmbDirectionSource.Picture = Application.CommandBars.GetImageMso("FileOpen", 32, 32)
    Set Me.cmbDirectionDestination.Picture = Application.CommandBars.GetImageMso("FileSave", 32, 32)
End Sub

