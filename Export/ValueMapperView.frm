VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValueMapperView 
   Caption         =   "Map Value Columns"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "ValueMapperView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ValueMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.Views"
Option Explicit
Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As ValueMapperViewModel
Attribute vm.VB_VarHelpID = -1
Private Const ICON_SIZE As Long = 16
Private msoImageList As ImageList

Private Type TFrmKeyMapper2View
    IsCancelled As Boolean
End Type

Private This As TFrmKeyMapper2View

Private Sub cmbBack_Click()
    vm.GoBack = True
    Me.Hide
End Sub

' ---
Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbAutoMap_Click()
    vm.Automap
End Sub

Private Sub cmbClearSearchLHS_Click()
    vm.LHSCriteria = vbNullString
End Sub

Private Sub cmbClearSearchRHS_Click()
    vm.RHSCriteria = vbNullString
End Sub

Private Sub cmbFinish_Click()
    Me.Hide
End Sub

Private Sub cmbMapRight_Click()
    vm.TryMap
End Sub

Private Sub cmbOptions_Click()
    Dim optionVM As TransferOptionsViewModel
    Set optionVM = New TransferOptionsViewModel
    optionVM.Flags = vm.Flags
    
    Dim View As IView
    Set View = New TransferOptionsView
    
    If View.ShowDialog(optionVM) Then
        vm.Flags = optionVM.Flags
    Else
        'Debug.Print "Cancelled"
    End If
End Sub

Private Sub cmbReset_Click()
    vm.Reset
End Sub

Private Sub cmbSelectNone_Click()
    vm.SelectNone
End Sub

Private Sub cmbSelectAll_Click()
    vm.SelectAll
End Sub

Private Sub cmbUnmapLeft_Click()
    vm.TryUnMap
End Sub

Private Sub chkShowMappedOnlyLHS_Click()
    vm.ShowMappedOnlyLHS = Me.chkShowMappedOnlyLHS.Value
End Sub

Private Sub chkShowMappedOnlyRHS_Click()
    vm.ShowMappedOnlyRHS = Me.chkShowMappedOnlyRHS.Value
End Sub

Private Sub lvLHS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    vm.TrySelectLHS Item
End Sub

Private Sub lvRHS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    vm.TrySelectRHS Item
End Sub

Private Sub lvRHS_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    vm.TryCheck Item
End Sub

Private Sub vm_CollectionChangedLHS()
    Dim current As String
    
    If Not Me.lvLHS.SelectedItem Is Nothing Then
        current = Me.lvLHS.SelectedItem.Key
    End If
    
    vm.LoadLHStoListView Me.lvLHS
    
    If current <> vbNullString Then
        If Not TryReselectListItem(Me.lvLHS, current) Then
            If Me.lvLHS.ListItems.Count > 0 Then
                Me.lvLHS.ListItems.Item(1).Selected = True
                lvLHS_ItemClick Me.lvLHS.ListItems.Item(1)
            End If
        End If
    End If
End Sub

Private Sub vm_CollectionChangedRHS()
    Dim current As String
    
    If Not Me.lvRHS.SelectedItem Is Nothing Then
        current = Me.lvRHS.SelectedItem.Key
    End If
    
    vm.LoadRHStoListView Me.lvRHS
    vm.UpdateRHStoListView Me.lvRHS
    
    If current <> vbNullString Then
        If Not TryReselectListItem(Me.lvRHS, current) Then
            If Me.lvRHS.ListItems.Count > 0 Then
                Me.lvRHS.ListItems.Item(1).Selected = True
                lvRHS_ItemClick Me.lvRHS.ListItems.Item(1)
            End If
        End If
    End If
End Sub

Private Function TryReselectListItem(ByVal lv As ListView, ByVal Key As String) As Boolean
    Dim i As Long
    For i = 1 To lv.ListItems.Count
        If lv.ListItems.Item(i).Key = Key Then
            lv.ListItems.Item(i).Selected = True
            TryReselectListItem = True
            Exit Function
        End If
    Next i
End Function

Private Sub vm_MappingChanged()
    vm.UpdateLHStoListView Me.lvLHS
    vm.UpdateRHStoListView Me.lvRHS
    
    Me.cmbReset.Enabled = vm.CanReset
    Me.cmbAutoMap.Enabled = vm.CanAutoMap
    Me.cmbSelectAll.Enabled = vm.CanSelectAll
    Me.cmbSelectNone.Enabled = vm.CanSelectNone
    
    vm_SelectionChanged
End Sub

Private Sub vm_SelectionChanged()
    Me.cmbMapRight.Enabled = vm.CanMapRight
    Me.cmbUnmapLeft.Enabled = vm.CanUnmapLeft
    
    Me.txtSearchLHS = vm.LHSCriteria
    Me.txtSearchRHS = vm.RHSCriteria
End Sub

Private Sub txtSearchLHS_Change()
    If IsNull(Me.txtSearchLHS) Then Exit Sub
    vm.LHSCriteria = Me.txtSearchLHS
End Sub

Private Sub txtSearchRHS_Change()
    If IsNull(Me.txtSearchRHS) Then Exit Sub
    vm.RHSCriteria = Me.txtSearchRHS
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
    Set vm = ViewModel
    This.IsCancelled = False
    
    Set msoImageList = New ImageList
    Set msoImageList = StandardImageList.GetMSOImageList(ICON_SIZE)
    
    Set Me.lvLHS.Icons = msoImageList
    Set Me.lvLHS.SmallIcons = msoImageList
    Set Me.lvRHS.Icons = msoImageList
    Set Me.lvRHS.SmallIcons = msoImageList
    
    LoadFromVM
    
    lvLHS_ItemClick Me.lvLHS.ListItems.Item(1)
    lvRHS_ItemClick Me.lvRHS.ListItems.Item(1)
    
    vm_MappingChanged
    
    Me.cmbClearSearchLHS.Picture = msoImageList.ListImages.Item("delete").Picture
    Me.cmbClearSearchRHS.Picture = msoImageList.ListImages.Item("delete").Picture
    
    vm.AutomapIfEmpty
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Public Sub LoadFromVM()
    vm.InitializeListView Me.lvLHS
    vm.InitializeListView Me.lvRHS, True
    vm.LoadLHStoListView Me.lvLHS
    vm.LoadRHStoListView Me.lvRHS
End Sub

