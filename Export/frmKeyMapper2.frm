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
Implements IView

Private WithEvents vm As clsKeyMapperViewModel
Attribute vm.VB_VarHelpID = -1
Private Const ICON_SIZE As Integer = 16
Private msoImageList As ImageList

Public DEBUG_EVENTS As Boolean

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

Private Sub cmbColumnLHS_Change()
    vm.TrySelectLHS Me.cmbColumnLHS
End Sub

Private Sub cmbColumnRHS_Change()
    vm.TrySelectRHS Me.cmbColumnRHS
End Sub

Private Sub cmbTableLHS_DropButtonClick()
    Dim result As ListObject
    If modSelectTable.TrySelectTable(result) = True Then
        Set vm.LHSTable = result
    End If
    Me.cmbColumnLHS.SetFocus
End Sub

Private Sub cmbTableRHS_DropButtonClick()
    Dim result As ListObject
    If modSelectTable.TrySelectTable(result) = True Then
        Set vm.RHSTable = result
    End If
    Me.cmbColumnRHS.SetFocus
End Sub

Private Sub ChangeTable(ByRef cmb As ComboBox, ByRef lo As ListObject)
    cmb.RemoveItem (0)
    cmb.AddItem lo.Name
    cmb = lo.Name
End Sub

Private Sub PopulateColumns(ByRef cmb As ComboBox, ByRef lo As ListObject)
    cmb.Clear
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        cmb.AddItem lc.Name
    Next lc
    cmb = cmb.List(0)
End Sub

Private Sub cmbNext_Click()
    Me.Hide
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
Private Function IView_ShowDialog(ByVal viewModel As IViewModel) As Boolean
    Set vm = viewModel
    this.IsCancelled = False
    
    Set msoImageList = New ImageList
    PopulateImageList msoImageList, ICON_SIZE
    
    InitializeTableCombobox Me.cmbTableLHS
    InitializeTableCombobox Me.cmbTableRHS
    
    LoadFromVM
    
    Me.Show
    
    IView_ShowDialog = Not this.IsCancelled
End Function

Private Sub InitializeTableCombobox(ByRef cmb As ComboBox)
    With cmb
        .Clear
        .AddItem "(No table selected)"
        .value = "(No table selected)"
    End With
End Sub

Public Sub LoadFromVM()
    If Not vm.LHSTable Is Nothing Then
        vm_PropertyChanged "LHSTable"
        vm_PropertyChanged "LHSColumns"
    End If
    If Not vm.RHSTable Is Nothing Then
        vm_PropertyChanged "RHSTable"
        vm_PropertyChanged "RHSColumns"
    End If
End Sub

Private Sub vm_PropertyChanged(ByVal propertyName As String)
    If DEBUG_EVENTS = True Then Debug.Print "Property changed: " & propertyName
    Select Case propertyName
        Case "LHSTable"
            ChangeTable Me.cmbTableLHS, vm.LHSTable
        Case "LHSColumns"
            PopulateColumns Me.cmbColumnLHS, vm.LHSTable
        Case "LHSKeyColumn"
            PopulateKeyPreview Me.lvQualityLHS, vm.LHSKeyColumn
        Case "RHSTable"
            ChangeTable Me.cmbTableRHS, vm.RHSTable
        Case "RHSColumns"
            PopulateColumns Me.cmbColumnRHS, vm.RHSTable
        Case "RHSKeyColumn"
            PopulateKeyPreview Me.lvQualityRHS, vm.RHSKeyColumn
    End Select
End Sub

Private Sub UpdateCheckButton()
    'Me.cmbCheck.Enabled = vm.CanCheck And vm.IsDirty
End Sub

Private Sub PopulateKeyPreview(ByRef lv As ListView, ByRef lc As ListColumn)
    Dim vm As clsColumnQualityViewModel
    Set vm = New clsColumnQualityViewModel
    Set vm.ListColumn = lc
    vm.InitializeListView lv
    vm.UpdateListView lv
End Sub
