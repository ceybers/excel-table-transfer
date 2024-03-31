VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MatchQualityView 
   Caption         =   "Match Quality Details"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "MatchQualityView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MatchQualityView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("MVVM.Views")
Option Explicit
Implements IView

Private Type TState
    ViewModel As KeyMapperViewModel
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
    UpdateControls
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    MatchQualityToListView.Initialize Me.lvLeftOnly
    MatchQualityToListView.Initialize Me.lvIntersection
    MatchQualityToListView.Initialize Me.lvRightOnly
    MatchQualityToTextBox.Initialize Me.txtMatchQuality
End Sub

Private Sub UpdateControls()
    MatchQualityToListView.Load Me.lvLeftOnly, This.ViewModel.MatchQuality.LeftOnly
    MatchQualityToListView.Load Me.lvIntersection, This.ViewModel.MatchQuality.Intersection
    MatchQualityToListView.Load Me.lvRightOnly, This.ViewModel.MatchQuality.RightOnly
    MatchQualityToTextBox.Load Me.txtMatchQuality, This.ViewModel.MatchQuality
End Sub
