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
    ViewModel As KeyQualityViewModel
    Result As ViewResult
End Type
Private This As TState

Private Sub cboCancel_Click()
    This.Result = vrCancel
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As ViewResult
    Set This.ViewModel = ViewModel
    
    InitializeControls
    UpdateListView
    UpdateButtons
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    KeyColumnToListView.Initialize Me.lvQuality
End Sub

Private Sub UpdateListView()
    KeyColumnToListView.Load Me.lvQuality, This.ViewModel.KeyColumn
End Sub

Private Sub UpdateButtons()
    Me.cboCancel.Enabled = True
End Sub


