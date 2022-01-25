VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValueMapperView 
   Caption         =   "Map Value Columns"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
   OleObjectBlob   =   "ValueMapperView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValueMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "ValueMapper"
Option Explicit
Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents Model As ValueMapperViewModel
Attribute Model.VB_VarHelpID = -1
Private Const ICON_SIZE As Integer = 16
Private msoImageList As ImageList

Public DEBUG_EVENTS As Boolean

Private Type TFrmKeyMapper2View
    IsCancelled As Boolean
End Type

Private This As TFrmKeyMapper2View

' ---
Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbAutoMap_Click()
    Model.Automap
End Sub

Private Sub cmbClearSearchLHS_Click()
   Model.LHSCriteria = vbNullString
End Sub

Private Sub cmbClearSearchRHS_Click()
    Model.RHSCriteria = vbNullString
End Sub

Private Sub cmbFinish_Click()
    Me.Hide
End Sub

Private Sub cmbMapRight_Click()
    Model.TryMap
End Sub

Private Sub cmbReset_Click()
    Model.Reset
End Sub

Private Sub cmbSelectNone_Click()
    Model.SelectNone
End Sub

Private Sub cmbSelectAll_Click()
    Model.SelectAll
End Sub

Private Sub cmbUnmapLeft_Click()
    Model.TryUnMap
End Sub

Private Sub lvLHS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Model.TrySelectLHS Item
End Sub

Private Sub lvRHS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Model.TrySelectRHS Item
End Sub

Private Sub lvRHS_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Model.TryCheck Item
End Sub

Private Sub Model_CollectionChangedLHS()
    Dim current As String
    
    If Not Me.lvLHS.SelectedItem Is Nothing Then
        current = Me.lvLHS.SelectedItem.key
    End If
    
    Model.LoadLHStoListView Me.lvLHS
    
    If current <> vbNullString Then
        If Not TryReselectListItem(Me.lvLHS, current) Then
            If Me.lvLHS.ListItems.Count > 0 Then
                Me.lvLHS.ListItems(1).Selected = True
                lvLHS_ItemClick Me.lvLHS.ListItems(1)
            End If
        End If
    End If
End Sub

Private Sub Model_CollectionChangedRHS()
    Dim current As String
    
    If Not Me.lvRHS.SelectedItem Is Nothing Then
        current = Me.lvRHS.SelectedItem.key
    End If
    
    Model.LoadRHStoListView Me.lvRHS
    Model.UpdateRHStoListView Me.lvRHS
    
    If current <> vbNullString Then
        If Not TryReselectListItem(Me.lvRHS, current) Then
            If Me.lvRHS.ListItems.Count > 0 Then
                Me.lvRHS.ListItems(1).Selected = True
                lvRHS_ItemClick Me.lvRHS.ListItems(1)
            End If
        End If
    End If
End Sub

Private Function TryReselectListItem(ByVal lv As ListView, ByVal key As String) As Boolean
    Dim i As Long
    For i = 1 To lv.ListItems.Count
        If lv.ListItems(i).key = key Then
            lv.ListItems(i).Selected = True
            TryReselectListItem = True
            Exit Function
        End If
    Next i
End Function

Private Sub Model_MappingChanged()
    Model.UpdateLHStoListView Me.lvLHS
    Model.UpdateRHStoListView Me.lvRHS
    
    Me.cmbReset.Enabled = Model.CanReset
    Me.cmbAutoMap.Enabled = Model.CanAutoMap
    Me.cmbSelectAll.Enabled = Model.CanSelectAll
    Me.cmbSelectNone.Enabled = Model.CanSelectNone
    
    Model_SelectionChanged
End Sub

Private Sub Model_SelectionChanged()
    Me.cmbMapRight.Enabled = Model.CanMapRight
    Me.cmbUnmapLeft.Enabled = Model.CanUnmapLeft
    
    Me.txtSearchLHS = Model.LHSCriteria
    Me.txtSearchRHS = Model.RHSCriteria
End Sub

Private Sub txtSearchLHS_Change()
    If IsNull(Me.txtSearchLHS) Then Exit Sub
    Model.LHSCriteria = Me.txtSearchLHS
End Sub

Private Sub txtSearchRHS_Change()
    If IsNull(Me.txtSearchRHS) Then Exit Sub
    Model.RHSCriteria = Me.txtSearchRHS
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
    Set Model = ViewModel
    This.IsCancelled = False
    
    Set msoImageList = New ImageList
    Set msoImageList = modStandardImageList.GetMSOImageList(ICON_SIZE)
    
    Set Me.lvLHS.Icons = msoImageList
    Set Me.lvLHS.SmallIcons = msoImageList
    Set Me.lvRHS.Icons = msoImageList
    Set Me.lvRHS.SmallIcons = msoImageList
    
    LoadFromVM
    
    lvLHS_ItemClick Me.lvLHS.ListItems(1)
    lvRHS_ItemClick Me.lvRHS.ListItems(1)
    
    Model_MappingChanged
    
    Me.cmbClearSearchLHS.Picture = msoImageList.ListImages("delete").Picture
    Me.cmbClearSearchRHS.Picture = msoImageList.ListImages("delete").Picture
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Public Sub LoadFromVM()
    Model.InitializeListView Me.lvLHS
    Model.InitializeListView Me.lvRHS, True
    Model.LoadLHStoListView Me.lvLHS
    Model.LoadRHStoListView Me.lvRHS
End Sub
