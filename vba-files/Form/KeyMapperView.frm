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


'@Folder "MVVM.KeyMapper"
Option Explicit
Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As KeyMapperViewModel
Attribute vm.VB_VarHelpID = -1
Private Const ICON_SIZE As Integer = 16
Private msoImageList As ImageList

Public DEBUG_EVENTS As Boolean

Private Const NO_TABLE_SELECTED As String = "(No table selected)"

Private Type TFrmKeyMapper2View
    SelectTableVM As SelectTableViewModel
    IsLoaded As Boolean
    IsCancelled As Boolean
    IsInitialLoad As Boolean
    IsUserChangingTable As Boolean
End Type

Private This As TFrmKeyMapper2View

' --- Controls
Private Sub chkAddNewKeys_Click()
    If Me.chkAddNewKeys.Value = True Then
        vm.AppendNewKeys = True
    ElseIf Me.chkAddNewKeys.Value = False Then
        vm.AppendNewKeys = False
    Else
        ' Tri state checkbox, do nothing?
    End If
End Sub

Private Sub chkLimitKeyCheck_Click()
    Me.txtLimitKeyValue.Enabled = Me.chkLimitKeyCheck.Value
End Sub

Private Sub chkRemoveOrphans_Click()
    vm.RemoveOrphanKeys = Me.chkRemoveOrphans.Value
End Sub

Private Sub cmbBack_Click()
    vm.GoBack = True
    Me.Hide
End Sub

Private Sub cmbBestMatch_Click()
    vm.TryGuess
End Sub

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbCheckKeys_Click()
    vm.UpdateMatch
End Sub

Private Sub cmbCheckQuality_Click()
    vm.UpdatePreviews
End Sub

Private Sub cmbColumnLHS_DropButtonClick()
    If Me.cmbColumnLHS = vm.LHSKeyColumn.Name Then
        Exit Sub
    End If
    
    vm.TrySelectLHS Me.cmbColumnLHS
    
    If This.IsInitialLoad = False Then
        vm.TryAutoMatch leftToRight:=True, Quiet:=False
    End If
End Sub

Private Sub cmbColumnRHS_DropButtonClick()
    If Me.cmbColumnRHS = vm.RHSKeyColumn.Name Then
        Exit Sub
    End If
    
    vm.TrySelectRHS Me.cmbColumnRHS
    
    If This.IsInitialLoad = False Then
        vm.TryAutoMatch leftToRight:=False, Quiet:=False
    End If
End Sub

Private Sub cmbHistory_Click()
    OnCancel
    'modMain.TransferTableFromHistory
End Sub

Private Sub cmbMatchLTR_Click()
    vm.TryAutoMatch True, True
End Sub

Private Sub cmbMatchRTL_Click()
    vm.TryAutoMatch False, True
End Sub

Private Sub cmbSwap_Click()
    vm.TrySwap
End Sub

Private Sub cmbTableLHS_DropButtonClick()
    If TrySelectTable(Nothing, This.SelectTableVM) Then
        Set vm.LHSTable = This.SelectTableVM.SelectedTable
    End If
    
    Me.cmbColumnLHS.SetFocus
End Sub

Private Sub cmbTableRHS_DropButtonClick()
    If TrySelectTable(Nothing, This.SelectTableVM) Then
        Set vm.RHSTable = This.SelectTableVM.SelectedTable
    End If
    
    Me.cmbColumnRHS.SetFocus
End Sub

Private Sub cmbNext_Click()
    Me.Hide
End Sub

' --- Subs
Private Sub ChangeTable(ByVal cmb As ComboBox, ByVal lo As ListObject)
    Debug.Assert cmb.ListCount > 0
    
    cmb.RemoveItem 0
    cmb.AddItem lo.Name
    cmb = lo.Name
End Sub

Private Sub PopulateColumns(ByVal cmb As ComboBox, ByVal lo As ListObject)
    Dim currentColumn As String
    currentColumn = cmb
    
    Dim canRememberColumn As Boolean
    canRememberColumn = False

    cmb.Clear

    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        cmb.AddItem lc.Name
        
        If lc.Name = currentColumn Then
            canRememberColumn = True
        End If
    Next lc
    
    If canRememberColumn Then
        cmb = currentColumn
    Else
        cmb = cmb.List(0)
    End If
End Sub

Private Sub txtLimitKeyValue_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.txtLimitKeyValue.Value) Then
        Cancel = True
    ElseIf CLng(Me.txtLimitKeyValue.Value) < 1 Then
        Cancel = True
    ElseIf CLng(Me.txtLimitKeyValue.Value) > 10000 Then
        Cancel = True
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

' ---
Private Function IView_ShowDialog(ByVal ViewModel As IViewModel) As Boolean
    PrintTime "KeyMapperView"
    DEBUG_EVENTS = False
    This.IsInitialLoad = True
    
    Set This.SelectTableVM = New SelectTableViewModel
    
    Set vm = ViewModel
    This.IsCancelled = False
    
    Set msoImageList = modStandardImageList.GetMSOImageList(ICON_SIZE)
    
    InitializeTableCombobox Me.cmbTableLHS
    InitializeTableCombobox Me.cmbTableRHS
    
    LoadTablesFromVM
    PrintTime "LoadTablesFromVM"
    LoadFlagsFromVM
    PrintTime "LoadFlagsFromVM"
    vm.TryGuess
    PrintTime "TryGuess"
    
    This.IsLoaded = True
    
    This.IsInitialLoad = False
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitializeTableCombobox(ByVal cmb As ComboBox)
    With cmb
        .Clear
        .AddItem NO_TABLE_SELECTED
        .Value = NO_TABLE_SELECTED
    End With
End Sub

Public Sub LoadTablesFromVM()
    If Not vm.LHSTable Is Nothing Then
        vm_PropertyChanged KeyMapperEvents.LHS_TABLE
        vm_PropertyChanged KeyMapperEvents.LHS_KEY_COLUMN
    End If
    
    If Not vm.LHSKeyColumn Is Nothing Then
        'vm_PropertyChanged KeyMapperEvents.LHS_KEY_COLUMN
    End If
    
    If Not vm.RHSTable Is Nothing Then
        vm_PropertyChanged KeyMapperEvents.RHS_TABLE
        vm_PropertyChanged KeyMapperEvents.RHS_KEY_COLUMN
    End If
    
    If Not vm.RHSKeyColumn Is Nothing Then
        'vm_PropertyChanged KeyMapperEvents.RHS_KEY_COLUMN
    End If
End Sub

Public Sub LoadFlagsFromVM()
    Me.chkAddNewKeys.Value = vm.AppendNewKeys
    Me.chkRemoveOrphans.Value = vm.RemoveOrphanKeys
    
    Me.chkAddNewKeys.Enabled = False
    Me.chkRemoveOrphans.Enabled = False
End Sub

Private Sub vm_MatchChanged()
    PopulateMatchSets
    UpdateCheckButton
    Me.cmbNext.SetFocus
End Sub

Private Sub vm_PreviewChanged()
    PopulateKeyPreview Me.lvQualityLHS, vm.LHSKeyColumn
    PopulateKeyPreview Me.lvQualityRHS, vm.RHSKeyColumn
    UpdateCheckButton
    Me.cmbCheckKeys.SetFocus
End Sub

Private Sub vm_PropertyChanged(ByVal propertyName As String)
    If DEBUG_EVENTS = True Then Debug.Print "Property changed: " & propertyName
    Select Case propertyName
        Case KeyMapperEvents.LHS_TABLE
            'this.IsInitialLoad = True
            ChangeTable Me.cmbTableLHS, vm.LHSTable
            PopulateColumns Me.cmbColumnLHS, vm.LHSTable
            TryAutoMatchAgain leftToRight:=False
            This.IsInitialLoad = False
            'this.IsUserChangingTable = False
        'Case KeyMapperEvents.LHS_COLUMNS
            
        Case KeyMapperEvents.LHS_KEY_COLUMN
            ResetQualityControls lhs:=True
        Case KeyMapperEvents.RHS_TABLE
            'this.IsInitialLoad = True
            ChangeTable Me.cmbTableRHS, vm.RHSTable
            PopulateColumns Me.cmbColumnRHS, vm.RHSTable
            TryAutoMatchAgain leftToRight:=True
            This.IsInitialLoad = False
            'this.IsUserChangingTable = False
        'Case KeyMapperEvents.RHS_COLUMNS
            
        Case KeyMapperEvents.RHS_KEY_COLUMN
            ResetQualityControls RHS:=True
    End Select
    
    Me.cmbSwap.Enabled = vm.CanSwap
    If Me.cmbColumnLHS = Me.cmbColumnRHS Then
        Me.cmbMatchLTR.Enabled = False
        Me.cmbMatchRTL.Enabled = False
    Else
        Me.cmbMatchLTR.Enabled = True
        Me.cmbMatchRTL.Enabled = True
    End If
End Sub

Private Sub TryAutoMatchAgain(ByVal leftToRight As Boolean)
    'If this.IsInitialLoad = True Then Exit Sub
    If vm.TryAutoMatch(leftToRight, True) = False Then
        vm.TryGuess
    End If
End Sub

Private Sub ResetQualityControls(Optional ByVal lhs As Boolean, Optional ByVal RHS As Boolean)
    If lhs Then
        Me.lvQualityLHS.ListItems.Clear
    End If

    If RHS Then
        Me.lvQualityRHS.ListItems.Clear
    End If

    Me.lvSetLHS.ListItems.Clear
    Me.lvSetInner.ListItems.Clear
    Me.lvSetRHS.ListItems.Clear

    UpdateCheckButton
    UpdateComboColumn
End Sub

Private Sub UpdateCheckButton()
    Me.cmbCheckQuality.Enabled = vm.CanCheck
    Me.cmbCheckKeys.Enabled = vm.CanMatch
    Me.cmbNext.Enabled = vm.CanContinue
End Sub

Private Sub UpdateComboColumn()
    If Not vm.LHSKeyColumn Is Nothing Then
        If Me.cmbColumnLHS <> vm.LHSKeyColumn.Name Then
            Me.cmbColumnLHS = vm.LHSKeyColumn.Name
        End If
    End If
    
    If Not vm.RHSKeyColumn Is Nothing Then
        If Me.cmbColumnRHS <> vm.RHSKeyColumn.Name Then
            Me.cmbColumnRHS = vm.RHSKeyColumn.Name
        End If
    End If

    If vm.CanCheck Then
        Me.cmbCheckQuality.SetFocus
    End If
End Sub

Private Sub PopulateKeyPreview(ByVal lv As ListView, ByVal lc As ListColumn)
    '@MemberAttribute VB_VarHelpID, -1
    Dim vm As ColumnQualityViewModel
    Set vm = New ColumnQualityViewModel
    Set vm.ListColumn = lc
    vm.InitializeListView lv
    vm.UpdateListView lv
    Set vm = Nothing
End Sub

Private Sub PopulateMatchSets()
    Dim comp As KeyColumnComparer
    Set comp = New KeyColumnComparer
    
    If Me.chkLimitKeyCheck.Value = True Then
        Set comp.SrcKeyColumn = KeyColumn.FromColumn(vm.LHSKeyColumn, False, Me.txtLimitKeyValue.Value)
        Set comp.DstKeyColumn = KeyColumn.FromColumn(vm.RHSKeyColumn, False, Me.txtLimitKeyValue.Value)
    Else
        Set comp.SrcKeyColumn = KeyColumn.FromColumn(vm.LHSKeyColumn)
        Set comp.DstKeyColumn = KeyColumn.FromColumn(vm.RHSKeyColumn)
    End If
    
    ' This maps the keys to their location in the other array
    ' It doesn't calculate LeftOny/Intersection/RightOnly -
    ' this is done on Set .LHS/.RHS
    'comp.Map
    
    CollectionToListView comp.SrcSetOnly, Me.lvSetLHS, "Additions"
    CollectionToListView comp.IntersectionSet, Me.lvSetInner, "Matches"
    CollectionToListView comp.DstSetOnly, Me.lvSetRHS, "Orphans"

    Me.chkAddNewKeys.Enabled = (Me.lvSetLHS.ListItems.Count > 0)
    Me.chkRemoveOrphans.Enabled = (Me.lvSetRHS.ListItems.Count > 0)

    Set comp = Nothing
End Sub

Private Sub CollectionToListView(ByVal coll As Collection, ByVal lv As ListView, ByVal header As String)
    With lv
        .view = lvwReport
        .Gridlines = True
        .ListItems.Clear
        .ColumnHeaders.Clear
    End With
    
    If coll Is Nothing Then
        lv.ColumnHeaders.Add text:=header & " (0)"
    Else
        Dim V As Variant
        For Each V In coll
            lv.ListItems.Add text:=V
        Next V
        lv.ColumnHeaders.Add text:=header & " (" & coll.Count & ")"
    End If
End Sub
