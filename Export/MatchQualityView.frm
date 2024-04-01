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
Implements IView3

Private Type TState
    Context As IAppContext
    ViewModel As KeyMapperViewModel
    Result As TtViewResult
End Type
Private This As TState

Private Property Get IView3_ViewModel() As Object
    Set IView3_ViewModel = This.ViewModel
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

Private Sub cboCancel_Click()
    This.Result = vrCancel
    Me.Hide
End Sub

Private Sub mpgMatchQuality_Change()
    If Me.mpgMatchQuality.Value = 1 Then
        Me.txtMatchQuality.SetFocus
        Me.txtMatchQuality.SelStart = 0
        Me.txtMatchQuality.SelLength = Len(Me.txtMatchQuality.Text)
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Sub IView3_Show()
    IView3_ShowDialog
End Sub
 
Private Sub IView3_Hide()
    Me.Hide
End Sub

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As KeyMapperViewModel) As IView3
    Dim Result As MatchQualityView
    Set Result = New MatchQualityView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView3_ShowDialog() As TtViewResult
    Set This.ViewModel = ViewModel
    
    InitializeControls
    UpdateControls
    
    Me.Show
    
    IView3_ShowDialog = This.Result
End Function

Private Sub InitializeControls()
    MatchQualityToListView.Initialize Me.lvLeftOnly
    MatchQualityToListView.Initialize Me.lvIntersection
    MatchQualityToListView.Initialize Me.lvRightOnly
    MatchQualityToTextBox.Initialize Me.txtMatchQuality
    
    Me.mpgMatchQuality.Value = 0
End Sub

Private Sub UpdateControls()
    MatchQualityToListView.Load Me.lvLeftOnly, This.ViewModel.MatchQuality.LeftOnly
    MatchQualityToListView.Load Me.lvIntersection, This.ViewModel.MatchQuality.Intersection
    MatchQualityToListView.Load Me.lvRightOnly, This.ViewModel.MatchQuality.RightOnly
    MatchQualityToTextBox.Load Me.txtMatchQuality, This.ViewModel.MatchQuality
End Sub
