VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TablePropView 
   Caption         =   "Table Properties"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   OleObjectBlob   =   "TablePropView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TablePropView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements iview

Private WithEvents mViewModel As TablePropViewModel

Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdCancel_Click()
    OnCancel
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub cmdActivateListObject_Click()
    mViewModel.DoActiveListObject
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set mViewModel = ViewModel
    This.IsCancelled = False
    
    Initalize
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub Initalize()
    UpdateControls
    InitializeLabelPictures
    mViewModel.ColumnProperties.LoadListView Me.lvStarredColumns
End Sub

Private Sub UpdateControls()
    Me.txtTableName = mViewModel.TableName
    Me.txtWorkSheetName = mViewModel.WorkSheetName
    Me.txtWorkBookName = mViewModel.WorkBookName
End Sub

Private Sub InitializeLabelPictures()
    Set Me.Label1.Picture = Application.CommandBars.GetImageMso("TableInsert", 32, 32)
    Set Me.Label7.Picture = Application.CommandBars.GetImageMso("RelationshipsEditRelationships", 32, 32)
    Set Me.Label8.Picture = Application.CommandBars.GetImageMso("FieldsMenu", 32, 32)
End Sub
