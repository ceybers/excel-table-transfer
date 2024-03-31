VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SourceOrDestinationView 
   Caption         =   "Source or Destination?"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "SourceOrDestinationView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "SourceOrDestinationView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@IgnoreModule HungarianNotation
'@Folder "MVVM.Views"
Option Explicit
Implements IView2

Private Const ICON_SIZE As Long = 64

Private Type TView
    ViewModel As SourceOrDestinationViewModel
    Result As ViewResult
End Type

Private This As TView

Private Sub cmbCancel_Click()
    This.Result = vrCancel
    Me.Hide
End Sub

Private Sub cmbDestination_Click()
    This.ViewModel.IsDestination = True
    This.Result = vrNext
    Me.Hide
End Sub

Private Sub cmbSource_Click()
    This.ViewModel.IsSource = True
    This.Result = vrNext
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Function IView2_ShowDialog(ByVal ViewModel As Object) As ViewResult
    Set This.ViewModel = ViewModel

    InitalizeControls
    Me.Show
    
    IView2_ShowDialog = This.Result
End Function

Private Sub InitalizeControls()
    ' TODO Replace these with frmPictures32 icons
    Set Me.Image1.Picture = Application.CommandBars.GetImageMso("CreateTable", ICON_SIZE, ICON_SIZE)
    Set Me.Image2.Picture = Application.CommandBars.GetImageMso("GroupPivotTableGroup", ICON_SIZE, ICON_SIZE)
    Set Me.Image3.Picture = Application.CommandBars.GetImageMso("CreateTable", ICON_SIZE, ICON_SIZE)
End Sub
