VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyQualityView 
   Caption         =   "Key Quality Details"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3225
   OleObjectBlob   =   "KeyQualityView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "KeyQualityView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("MVVM.Views")
Option Explicit
Implements IView

Private Type TState
    Context As IAppContext
    ViewModel As KeyQualityViewModel
    Result As TtViewResult
End Type
Private This As TState

Private Sub cboCancel_Click()
    This.Result = vrCancel
    Me.Hide
End Sub

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As KeyQualityViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As KeyQualityViewModel)
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
        This.Result = vrCancel
    End If
End Sub

Private Sub IView_Show()
    IView_ShowDialog
End Sub
 
Private Sub IView_Hide()
    Me.Hide
End Sub

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As KeyQualityViewModel) As IView
    Dim Result As KeyQualityView
    Set Result = New KeyQualityView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView_ShowDialog() As TtViewResult
    BindControls
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "KeyColumn", Me.lvQuality, "ListItems", OneWayBinding, KeyColumnToListView
    End With
End Sub
