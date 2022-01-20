VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKeyMapper2 
   Caption         =   "Key Column Mapper"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "frmKeyMapper2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKeyMapper2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "KeyMapper"
Option Explicit

Private WithEvents vm As clsKeyMapperViewModel
Attribute vm.VB_VarHelpID = -1
Private Const ICON_SIZE As Integer = 16
Private msoImageList As ImageList

Private Type TFrmKeyMapper2View
    IsCancelled As Boolean
End Type

Private this As TFrmKeyMapper2View

' ---
Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbCheck_Click()
    vm.DoCheck
End Sub

Private Sub cmbLHSTableSelect_Click()
    Dim selTblvm As clsSelectTableViewModel
    Set selTblvm = New clsSelectTableViewModel
    With New frmSelectTable
        If .ShowDialog(selTblvm) Then
            If Not selTblvm.SelectedTable Is Nothing Then
                Set vm.LHSTable = selTblvm.SelectedTable
            End If
        End If
    End With
End Sub

Private Sub cmbRHSTableSelect_Click()
    Dim selTblvm As clsSelectTableViewModel
    Set selTblvm = New clsSelectTableViewModel
    With New frmSelectTable
        If .ShowDialog(selTblvm) Then
            If Not selTblvm.SelectedTable Is Nothing Then
                Set vm.RHSTable = selTblvm.SelectedTable
            End If
        End If
    End With
End Sub

Private Function TrySelectTable() As ListObject
    
End Function

Private Sub cmbOK_Click()
    Me.Hide
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox1_DropButtonClick()
    cmbLHSTableSelect_Click
    Me.txtLHSTable.Locked = False
    Me.txtLHSTable.SetFocus
    Me.ComboBox1.AddItem Me.txtLHSTable.text
    Me.ComboBox1.Value = Me.txtLHSTable.text
End Sub

Private Sub imgcmbLHS_Click()
    If Not Me.imgcmbLHS.SelectedItem Is Nothing Then
        vm.TrySelectLHS Me.imgcmbLHS.SelectedItem.key
    End If
End Sub

Private Sub imgcmbRHS_Click()
    If Not Me.imgcmbRHS.SelectedItem Is Nothing Then
        vm.TrySelectRHS Me.imgcmbRHS.SelectedItem.key
    End If
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

' ---
Private Sub vm_CollectionChanged()
    LoadTreeview
End Sub

Private Sub vm_ItemSelected()
    If vm.SelectedTable Is Nothing Then
        Me.cmbOK.Enabled = False
    Else
        Me.cmbOK.Enabled = True
    End If
End Sub

Public Function ShowDialog(ByVal viewModel As Object) As Boolean
    Set vm = viewModel
    this.IsCancelled = False
    Me.cmbOK.Enabled = False
    
    Set msoImageList = New ImageList
    PopulateImageList msoImageList, ICON_SIZE
    Me.imgcmbLHS.ImageList = msoImageList
    Me.imgcmbRHS.ImageList = msoImageList
    
    Set Me.cmbLHSTableSelect.Picture = msoImageList.ListImages("FindText").Picture
    Set Me.cmbRHSTableSelect.Picture = msoImageList.ListImages("FindText").Picture
    
    vm_PropertyChanged "LHSTable"
    vm_PropertyChanged "LHSColumns"
    vm_PropertyChanged "RHSTable"
    vm_PropertyChanged "RHSColumns"
    
    imgcmbLHS_Click
    imgcmbRHS_Click
    
    UpdateCheckButton
    
    Show
    
    ShowDialog = Not this.IsCancelled
End Function

Private Sub PopulateImageCombo(ByRef imgCmb As ImageCombo, ByRef coll As Collection)
    Dim lc As ListColumn
    
    imgCmb.ComboItems.Clear
    
    For Each lc In coll
        imgCmb.ComboItems.Add key:=lc.Name, text:=lc.Name, image:="col", SelImage:="AdpPrimaryKey"
    Next lc
    
    If imgCmb.ComboItems.count > 0 Then
        imgCmb.ComboItems(1).Selected = True
    End If
End Sub

Private Sub vm_CheckCompleted()
    Me.lbIntersect.AddItem "Hi"
    UpdateCheckButton
End Sub

Private Sub vm_PropertyChanged(ByVal propertyName As String)
    Debug.Print "Property changed: " & propertyName
    Select Case propertyName
        Case "LHSTable"
            Me.txtLHSTable.text = vm.LHSTable.Name
        Case "LHSColumns"
            PopulateImageCombo Me.imgcmbLHS, vm.LHSColumns
        Case "LHSKeyColumn"
            UpdateCheckButton
        Case "RHSTable"
            Me.txtRHSTable.text = vm.RHSTable.Name
        Case "RHSColumns"
            PopulateImageCombo Me.imgcmbRHS, vm.RHSColumns
        Case "RHSKeyColumn"
            UpdateCheckButton
    End Select
End Sub

Private Sub UpdateCheckButton()
    Me.cmbCheck.Enabled = vm.CanCheck And vm.IsDirty
End Sub
