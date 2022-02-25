VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SourceOrDestinationView 
   Caption         =   "Source or Destination?"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4440
   OleObjectBlob   =   "SourceOrDestinationView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SourceOrDestinationView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SourceOrDestination"
Option Explicit
Implements IView

Private vm As SourceOrDestinationViewModel
Attribute vm.VB_VarHelpID = -1
Private Const ICON_SIZE As Long = 64

Private Type TView
    ' Context as MVVM.IAppContext
    'ViewModel As SelectTableViewModel
    IsCancelled As Boolean
End Type

Private this As TView

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbDestination_Click()
    vm.IsDestination = True
    Me.Hide
End Sub

Private Sub cmbSource_Click()
    vm.IsSource = True
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Set Me.Image1.Picture = Application.CommandBars.GetImageMso("CreateTable", ICON_SIZE, ICON_SIZE)
    Set Me.Image2.Picture = Application.CommandBars.GetImageMso("GroupPivotTableGroup", ICON_SIZE, ICON_SIZE)
    Set Me.Image3.Picture = Application.CommandBars.GetImageMso("CreateTable", ICON_SIZE, ICON_SIZE)
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

Private Function IView_ShowDialog(ByVal ViewModel As IViewModel) As Boolean
    Set vm = ViewModel

    Me.Show
    IView_ShowDialog = Not this.IsCancelled
End Function
