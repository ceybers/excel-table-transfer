VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmColumnQuality 
   Caption         =   "UserForm1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "frmColumnQuality.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmColumnQuality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("ColumnQuality")
Option Explicit
Implements IView
 
Private Type TView
    IsCancelled As Boolean
    Model As clsColumnQualityViewModel
End Type

Private this As TView
 
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
    this.IsCancelled = True
    Me.Hide
End Sub
 
Private Function IView_ShowDialog(ByVal viewModel As IViewModel) As Boolean
    Debug.Assert Not viewModel Is Nothing
    Set this.Model = viewModel
    
    this.Model.InitializeListView Me.ListView1
    this.Model.UpdateListView Me.ListView1
    
    Show
    IView_ShowDialog = Not this.IsCancelled
End Function
