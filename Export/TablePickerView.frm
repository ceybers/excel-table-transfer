VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TablePickerView 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "TablePickerView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "TablePickerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.Views"
Option Explicit
Implements IView

Private Type TState
    Context As IAppContext
    ViewModel As TablePickerViewModel
    Result As TtViewResult
End Type
Private This As TState

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As TablePickerViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As TablePickerViewModel)
    Set This.ViewModel = vNewValue
End Property

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set This.Context = vNewValue
End Property

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

Private Sub cboSelSrc_Click()
    This.ViewModel.PickSelectedTable ttSource
    UpdateListViewLHS
    TryAutoFocusNext
End Sub

Private Sub cboSelDst_Click()
    This.ViewModel.PickSelectedTable ttDestination
    UpdateListViewLHS
    TryAutoFocusNext
End Sub

Private Sub tvTables_NodeClick(ByVal Node As MSComctlLib.Node)
    This.ViewModel.TrySelect Node.Key
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Sub lblHeaderIcon_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Private Sub IView_Show()
    IView_ShowDialog
End Sub
 
Private Sub IView_Hide()
    Me.Hide
End Sub

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As TablePickerViewModel) As IView
    Dim Result As TablePickerView
    Set Result = New TablePickerView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView_ShowDialog() As TtViewResult
    Set This.ViewModel = ViewModel
    Me.lblHeaderText.Caption = HDR_TXT_TABLE_PICKER
    
    InitializeControls
    UpdateListViewLHS
    
    BindControls
    
    Me.Show vbModal
    
    IView_ShowDialog = This.Result
End Function

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath This.ViewModel, "SourceWorkbookName", Me.txtSrcWorkbook, "Value", OneWayBinding
        .BindPropertyPath This.ViewModel, "SourceTableName", Me.txtSrcTable, "Value", OneWayBinding
        .BindPropertyPath This.ViewModel, "DestinationWorkbookName", Me.txtDstWorkbook, "Value", OneWayBinding
        .BindPropertyPath This.ViewModel, "DestinationTableName", Me.txtDstTable, "Value", OneWayBinding
        
        .BindPropertyPath This.ViewModel, "CanPickSelected", Me.cboSelSrc, "Enabled", OneWayBinding
        .BindPropertyPath This.ViewModel, "CanPickSelected", Me.cboSelDst, "Enabled", OneWayBinding
        
        .BindPropertyPath This.ViewModel, "CanNext", Me.cboNext, "Enabled", OneWayBinding
    End With
End Sub

Private Sub InitializeControls()
    TablePickerToTreeView.Initialize Me.tvTables
End Sub

Private Sub UpdateListViewLHS()
    TablePickerToTreeView.Load Me.tvTables, This.ViewModel
End Sub

Private Sub TryAutoFocusNext()
    If Me.cboNext.Enabled = True Then
        Me.cboNext.SetFocus
    End If
End Sub
