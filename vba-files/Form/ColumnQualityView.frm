VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColumnQualityView 
   Caption         =   "UserForm1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "ColumnQualityView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColumnQualityView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "MVVM.ColumnQuality"
Option Explicit
Implements IView
 
Private Type TView
    IsCancelled As Boolean
    Model As ColumnQualityViewModel
End Type

Private This As TView
 
Private Sub OkButton_Click()
    Me.Hide
End Sub
 
Private Sub CancelButton_Click()
    OnCancel
End Sub
 
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
 
Private Function IView_ShowDialog(ByVal ViewModel As IViewModel) As Boolean
    Debug.Assert Not ViewModel Is Nothing
    Set This.Model = ViewModel
    
    This.Model.InitializeListView Me.ListView1
    This.Model.UpdateListView Me.ListView1
    
    Me.Show
    IView_ShowDialog = Not This.IsCancelled
End Function
