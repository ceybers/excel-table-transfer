VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeltasPreviewView 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "DeltasPreviewView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "DeltasPreviewView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule HungarianNotation
'@Folder "MVVM.Views"
Option Explicit
Implements IView

Private Type TState
    Context As IAppContext
    ViewModel As DeltasPreviewViewModel
    Result As TtViewResult
End Type
Private This As TState

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As DeltasPreviewViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As DeltasPreviewViewModel)
    Set This.ViewModel = vNewValue
End Property

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set This.Context = vNewValue
End Property

Private Sub cmbBack_Click()
    This.Result = vrBack
    Me.Hide
End Sub

Private Sub cmbCancel_Click()
    This.Result = vrCancel
    Me.Hide
End Sub

Private Sub cmbFinish_Click()
    This.Result = vrFinish
    Me.Hide
End Sub

Private Sub cmbNext_Click()
    This.Result = vrNext
    Me.Hide
End Sub

Private Sub cmbStart_Click()
    This.Result = vrStart
    Me.Hide
End Sub

Private Sub cmbShowAll_Click()
    DoShowAll
End Sub

Private Sub lvKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
    DoSelectKey Item.Text ' TODO Create a proper Model and VM
End Sub

Private Sub lvFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    DoSelectField Item.Text ' TODO Create a proper Model and VM
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

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As DeltasPreviewViewModel) As IView
    Dim Result As DeltasPreviewView
    Set Result = New DeltasPreviewView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView_ShowDialog() As TtViewResult
    Me.lblHeaderText.Caption = HDR_TXT_DELTAS_PREVIEW
    
    BindControls
    UpdateNoChanges
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "Keys", Me.lvKeys, "ListItems", TwoWayBinding, TransferDeltasToListView
        .BindPropertyPath ViewModel, "Fields", Me.lvFields, "ListItems", TwoWayBinding, TransferDeltasToListView
        .BindPropertyPath ViewModel, "Deltas", Me.lvDeltas, "ListItems", OneWayBinding, TransferDeltasToListView
    End With
End Sub

Private Sub UpdateNoChanges()
    If This.ViewModel.CanFinish = False Then
        Me.lblHeaderText.Caption = DELTAS_PREVIEW_NO_RESULTS
    End If
End Sub

Private Sub DoShowAll()
    This.ViewModel.DoShowAll
End Sub

Private Sub DoSelectKey(ByVal Key As String)
    If Key = SELECT_ALL Then
        This.ViewModel.TrySelectKey vbNullString
    Else
        This.ViewModel.TrySelectKey Key
    End If
End Sub

Private Sub DoSelectField(ByVal Field As String)
    If Field = SELECT_ALL Then
        This.ViewModel.TrySelectField vbNullString
    Else
        This.ViewModel.TrySelectField Field
    End If
End Sub
