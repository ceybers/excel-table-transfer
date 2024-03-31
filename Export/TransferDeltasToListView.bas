Attribute VB_Name = "TransferDeltasToListView"
'@Folder "MVVM2.ValueConverters"
Option Explicit

Private Const KEY_HEADER As String = "Keys"
Private Const FIELD_HEADER As String = "Fields"
Public Const SELECT_ALL As String = "(Select all)"
Private Const SELECT_ALL_COLOR As String = 8421504 'RGB(128, 128, 128)

Public Enum tdType
    tdKeyMember
    tdField
    tdDelta
End Enum

Public Sub Initialize(ByVal ListView As MSComctlLib.ListView, ByVal Member As tdType)
    With ListView
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
        .HideSelection = False
    End With
    
    SetColumnHeaders ListView, Member
End Sub

Private Sub SetColumnHeaders(ByVal ListView As MSComctlLib.ListView, ByVal Member As tdType)
    With ListView
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="#", Width:=ListView.Width - 16
    End With
    
    If Member = tdDelta Then
        ListView.ColumnHeaders.Item(1).Text = "Before"
        ListView.ColumnHeaders.Add Text:="After"
        ListView.ColumnHeaders.Item(1).Width = (ListView.Width - 16) / 2
        ListView.ColumnHeaders.Item(2).Width = (ListView.Width - 16) / 2
    End If
End Sub

Public Sub Load(ByVal ListView As MSComctlLib.ListView, ByVal TransferDeltasViewModel As TransferDeltasViewModel, _
    ByVal Member As tdType)
    ListView.ListItems.Clear
    
    'If TransferDeltasViewModel Is Nothing Then Exit Sub
    Debug.Assert Not TransferDeltasViewModel Is Nothing
    Debug.Assert Not TransferDeltasViewModel.Deltas Is Nothing
    
    Dim Item As Variant
    If Member = tdDelta Then
        Dim TransferDelta As TransferDelta
        For Each TransferDelta In TransferDeltasViewModel.Deltas
            AddItemTransferDelta ListView, TransferDelta
        Next TransferDelta
    ElseIf Member = tdKeyMember Then
        AddSelectAll ListView
        For Each Item In TransferDeltasViewModel.Keys
            AddItem ListView, Item
        Next Item
        UpdateHeader ListView, Member, (TransferDeltasViewModel.Keys.Count - 1)
    ElseIf Member = tdField Then
        AddSelectAll ListView
        For Each Item In TransferDeltasViewModel.Fields
            AddItem ListView, Item
        Next Item
        UpdateHeader ListView, Member, (TransferDeltasViewModel.Fields.Count - 1)
    End If
End Sub

Private Sub AddItemTransferDelta(ByVal ListView As MSComctlLib.ListView, ByVal TransferDelta As TransferDelta)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=TransferDelta.ValueBefore)
    ListItem.ListSubItems.Add Text:=TransferDelta.ValueAfter
End Sub

Private Sub AddItem(ByVal ListView As MSComctlLib.ListView, ByVal Text As String)
    ListView.ListItems.Add Text:=Text
End Sub

Private Sub UpdateHeader(ByVal ListView As MSComctlLib.ListView, ByVal Member As tdType, ByVal Count As Long)
    Dim HeaderText As String
    
    Select Case Member
        Case tdType.tdKeyMember
            HeaderText = KEY_HEADER
        Case tdType.tdField
            HeaderText = FIELD_HEADER
    End Select
    
    If Count = -1 Then
        ListView.ColumnHeaders.Item(1).Text = vbNullString
        ListView.ListItems.Clear
    Else
        ListView.ColumnHeaders.Item(1).Text = HeaderText & " (" & CStr(Count + 0) & ")"
    End If
End Sub

Private Sub AddSelectAll(ByVal ListView As MSComctlLib.ListView)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=SELECT_ALL)
    ListItem.ForeColor = SELECT_ALL_COLOR
    ListItem.Key = SELECT_ALL
End Sub

