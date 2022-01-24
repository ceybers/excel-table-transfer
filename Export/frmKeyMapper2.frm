VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKeyMapper2 
   Caption         =   "Key Column Mapper"
   ClientHeight    =   8445.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
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

Private Sub cmbTableLHS_Click()
    Dim result As ListObject
    If modSelectTable.TrySelectTable(result) = True Then
        Set vm.LHSTable = result
        ' TODO this does not go here
        'Me.cmbTableLHS.text = result.Name
    End If
End Sub

Private Sub cmbTableRHS_Click()
    Dim result As ListObject
    If modSelectTable.TrySelectTable(result) = True Then
        Set vm.RHSTable = result
        ' TODO this does not go here
        'Me.cmbTableRHS.text = result.Name
    End If
End Sub

Private Sub cmbTableRHS_DropButtonClick()
    If Me.cmbTableRHS.ListCount > 0 Then
        cmbTableRHS_Click
        Me.cmbCancel.SetFocus
        Me.cmbTableRHS.Clear
        ' ???
        Me.cmbTableRHS.AddItem vm.LHSTable.Name
        Me.cmbTableRHS.value = vm.LHSTable.Name
        Me.cmbTableRHS.SetFocus
    Else
        Me.cmbTableRHS.AddItem vm.LHSTable.Name
        Me.cmbTableRHS.value = vm.LHSTable.Name
        Me.cmbTableRHS.SetFocus
    End If
End Sub

Private Function TrySelectTable() As ListObject
    
End Function

Private Sub cmbOK_Click()
    Me.Hide
End Sub

Private Sub ComboBox1_DropButtonClick()
    'cmbLHSTableSelect_Click
    'Me.txtLHSTable.Locked = False
    'Me.txtLHSTable.SetFocus
    'Me.ComboBox1.AddItem Me.txtLHSTable.text
    'Me.ComboBox1.value = Me.txtLHSTable.text
End Sub

Private Sub imgcmbLHS_Click()
    'If Not Me.imgcmbLHS.SelectedItem Is Nothing Then
    '    vm.TrySelectLHS Me.imgcmbLHS.SelectedItem.key
    'End If
End Sub

Private Sub imgcmbRHS_Click()
    'If Not Me.imgcmbRHS.SelectedItem Is Nothing Then
    '    vm.TrySelectRHS Me.imgcmbRHS.SelectedItem.key
    'End If
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
Public Function ShowDialog(ByVal viewModel As Object) As Boolean
    Set vm = viewModel
    this.IsCancelled = False
    'Me.cmbOK.Enabled = False
    
    Set msoImageList = New ImageList
    PopulateImageList msoImageList, ICON_SIZE
    'Me.imgcmbLHS.ImageList = msoImageList
    'Me.imgcmbRHS.ImageList = msoImageList
    
    'Set Me.cmbLHSTableSelect.Picture = msoImageList.ListImages("FindText").Picture
    'Set Me.cmbRHSTableSelect.Picture = msoImageList.ListImages("FindText").Picture
    
    'vm_PropertyChanged "LHSTable"
    'vm_PropertyChanged "LHSColumns"
    'vm_PropertyChanged "RHSTable"
    'vm_PropertyChanged "RHSColumns"
    
    'imgcmbLHS_Click
    'imgcmbRHS_Click
    
    'UpdateCheckButton
    
    Me.cmbTableRHS.Clear
    Me.cmbTableRHS.AddItem "(No table selected)"
    Me.cmbTableRHS.value = "(No table selected)"
    
    zzzdowork
    
    Show
    
    ShowDialog = Not this.IsCancelled
End Function

Public Sub zzzdowork()
    Dim vm As clsColumnQualityViewModel
    Set vm = New clsColumnQualityViewModel
    Set vm.ListColumn = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    vm.InitializeListView Me.lvQualityLHS
    vm.UpdateListView Me.lvQualityLHS
    vm.InitializeListView Me.lvQualityRHS
    vm.UpdateListView Me.lvQualityRHS
    
    Me.cmbTableLHS.AddItem ("Table1")
    Me.cmbTableLHS.value = "Table1"
    Me.cmbColumnLHS.AddItem ("Column1")
    Me.cmbColumnLHS.value = "Column1"
End Sub

Private Sub PopulateImageCombo(ByRef imgCmb As ImageCombo, ByRef coll As Collection)
    Dim lc As ListColumn
    
    imgCmb.ComboItems.Clear
    
    For Each lc In coll
        imgCmb.ComboItems.Add key:=lc.Name, text:=lc.Name, image:="col", SelImage:="AdpPrimaryKey"
    Next lc
    
    If imgCmb.ComboItems.Count > 0 Then
        imgCmb.ComboItems(1).Selected = True
    End If
End Sub

Private Sub vm_CheckCompleted()
    'Me.lbIntersect.AddItem "Hi"
    UpdateCheckButton
End Sub

Private Sub vm_PropertyChanged(ByVal propertyName As String)
    Debug.Print "Property changed: " & propertyName
    Select Case propertyName
        Case "LHSTable"
            'Me.txtLHSTable.text = vm.LHSTable.Name
        Case "LHSColumns"
            'PopulateImageCombo Me.imgcmbLHS, vm.LHSColumns
        Case "LHSKeyColumn"
            'UpdateCheckButton
        Case "RHSTable"
            'Me.txtRHSTable.text = vm.RHSTable.Name
        Case "RHSColumns"
            'PopulateImageCombo Me.imgcmbRHS, vm.RHSColumns
        Case "RHSKeyColumn"
            'UpdateCheckButton
    End Select
End Sub

Private Sub UpdateCheckButton()
    'Me.cmbCheck.Enabled = vm.CanCheck And vm.IsDirty
End Sub
