VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValueMapperView 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "ValueMapperView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ValueMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.Views"
Option Explicit
Implements IView

Private Type TState
    Context As IAppContext
    ViewModel As ValueMapperViewModel
    Result As TtViewResult
End Type
Private This As TState

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
End Sub

Private Sub lvDstValues_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.Destination.TrySelect Item.Key
End Sub

Private Sub cboMapAdd_Click()
    This.ViewModel.DoMapAdd
End Sub

Private Sub cboMapRemove_Click()
    This.ViewModel.DoMapRemove
End Sub

Private Sub cboMapAll_Click()
    This.ViewModel.DoAutoMap
    If Me.cboNext.Enabled Then Me.cboNext.SetFocus
End Sub

Private Sub cboMapReset_Click()
    This.ViewModel.DoRemoveAll
    Me.cboMapAll.SetFocus
End Sub

Private Sub IView_Show()
    IView_ShowDialog
End Sub
 
Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Sub lblHeaderIcon_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As ValueMapperViewModel) As IView
    Dim Result As ValueMapperView
    Set Result = New ValueMapperView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView_ShowDialog() As TtViewResult
    Me.lblHeaderText.Caption = HDR_TXT_VALUE_MAPPER
    
    BindControls
    Me.cboMapAll.SetFocus
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Public Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "Source", Me.lvSrcValues, "ListItems", TwoWayBinding, ValueColumnsToListView
        .BindPropertyPath ViewModel, "Destination", Me.lvDstValues, "ListItems", TwoWayBinding, ValueColumnsToListView
        .BindPropertyPath ViewModel, "Mapped", Me.lvMappedValues, "ListItems", TwoWayBinding, ColumnPairsToListView
        
        .BindPropertyPath ViewModel, "CanNext", Me.cboNext, "Enabled", OneWayBinding
        .BindPropertyPath ViewModel, "CanMapAdd", Me.cboMapAdd, "Enabled", OneWayBinding
        .BindPropertyPath ViewModel, "CanMapRemove", Me.cboMapRemove, "Enabled", OneWayBinding
        .BindPropertyPath ViewModel, "CanRemoveAll", Me.cboMapReset, "Enabled", OneWayBinding
    End With
End Sub
