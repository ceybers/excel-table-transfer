VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyMapperView 
   Caption         =   "Key Column Mapper"
   ClientHeight    =   8445.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   OleObjectBlob   =   "KeyMapperView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KeyMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "KeyMapper"
Option Explicit
Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As KeyMapperViewModel
Attribute vm.VB_VarHelpID = -1
Private Const ICON_SIZE As Integer = 16
Private msoImageList As ImageList

Public DEBUG_EVENTS As Boolean

Private Type TFrmKeyMapper2View
    IsCancelled As Boolean
End Type

Private this As TFrmKeyMapper2View

Private Sub cmbBack_Click()
    MsgBox "NYI"
End Sub

' ---
Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbCheckKeys_Click()
    vm.UpdateMatch
End Sub

Private Sub cmbCheckQuality_Click()
    vm.UpdatePreviews
End Sub

Private Sub cmbColumnLHS_Change()
    vm.TrySelectLHS Me.cmbColumnLHS
End Sub

Private Sub cmbColumnRHS_Change()
    vm.TrySelectRHS Me.cmbColumnRHS
End Sub

Private Sub cmbTableLHS_DropButtonClick()
    Dim result As ListObject
    If TestSelectTable.TrySelectTable(result) = True Then
        Set vm.LHSTable = result
    End If
    Me.cmbColumnLHS.SetFocus
End Sub

Private Sub cmbTableRHS_DropButtonClick()
    Dim result As ListObject
    If TestSelectTable.TrySelectTable(result) = True Then
        Set vm.RHSTable = result
    End If
    Me.cmbColumnRHS.SetFocus
End Sub

Private Sub ChangeTable(ByVal cmb As ComboBox, ByVal lo As ListObject)
    cmb.RemoveItem (0)
    cmb.AddItem lo.Name
    cmb = lo.Name
End Sub

Private Sub PopulateColumns(ByVal cmb As ComboBox, ByVal lo As ListObject)
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
    
    Set msoImageList = modStandardImageList.GetMSOImageList(ICON_SIZE)
    
    InitializeTableCombobox Me.cmbTableLHS
    InitializeTableCombobox Me.cmbTableRHS
    
    LoadFromVM
    
    Me.Show
    
    IView_ShowDialog = Not this.IsCancelled
End Function

Private Sub InitializeTableCombobox(ByVal cmb As ComboBox)
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

Private Sub vm_MatchChanged()
    PopulateMatchSets
    UpdateCheckButton
End Sub

Private Sub vm_PreviewChanged()
    PopulateKeyPreview Me.lvQualityLHS, vm.LHSKeyColumn
    PopulateKeyPreview Me.lvQualityRHS, vm.RHSKeyColumn
    UpdateCheckButton
End Sub

Private Sub vm_PropertyChanged(ByVal propertyName As String)
    If DEBUG_EVENTS = True Then Debug.Print "Property changed: " & propertyName
    Select Case propertyName
        Case "LHSTable"
            ChangeTable Me.cmbTableLHS, vm.LHSTable
        Case "LHSColumns"
            PopulateColumns Me.cmbColumnLHS, vm.LHSTable
        Case "LHSKeyColumn"
            Me.lvQualityLHS.ListItems.Clear
            Me.lvSetLHS.ListItems.Clear
            Me.lvSetRHS.ListItems.Clear
            Me.lvSetInner.ListItems.Clear
            UpdateCheckButton
        Case "RHSTable"
            ChangeTable Me.cmbTableRHS, vm.RHSTable
        Case "RHSColumns"
            PopulateColumns Me.cmbColumnRHS, vm.RHSTable
        Case "RHSKeyColumn"
            Me.lvQualityRHS.ListItems.Clear
            Me.lvSetLHS.ListItems.Clear
            Me.lvSetRHS.ListItems.Clear
            Me.lvSetInner.ListItems.Clear
            UpdateCheckButton
           
    End Select
End Sub

Private Sub UpdateCheckButton()
    Debug.Print "UpdateCheckButton"
    Me.cmbCheckQuality.Enabled = vm.CanCheck
    Me.cmbCheckKeys.Enabled = vm.CanMatch
    Me.cmbNext.Enabled = vm.CanContinue
End Sub

Private Sub PopulateKeyPreview(ByVal lv As ListView, ByVal lc As ListColumn)
    '@MemberAttribute VB_VarHelpID, -1
    Dim vm As ColumnQualityViewModel
    Set vm = New ColumnQualityViewModel
    Set vm.ListColumn = lc
    vm.InitializeListView lv
    vm.UpdateListView lv
End Sub

Private Sub PopulateMatchSets()
    Dim comp As KeyColumnComparer
    Set comp = New KeyColumnComparer
    Set comp.LHS = KeyColumn.FromColumn(vm.LHSKeyColumn)
    Set comp.RHS = KeyColumn.FromColumn(vm.RHSKeyColumn)
    comp.Map
    
    CollectionToListView comp.LeftOnly, Me.lvSetLHS
    CollectionToListView comp.Intersection, Me.lvSetInner
    CollectionToListView comp.RightOnly, Me.lvSetRHS
End Sub

Private Sub CollectionToListView(ByVal coll As Collection, ByVal lv As ListView)
    lv.view = lvwReport
    lv.Gridlines = True
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    Dim v As Variant
    For Each v In coll
        lv.ListItems.Add text:=v
    Next v
    lv.ColumnHeaders.Add text:="Key Items (" & coll.Count & ")"
End Sub
