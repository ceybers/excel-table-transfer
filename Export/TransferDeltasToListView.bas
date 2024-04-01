Attribute VB_Name = "TransferDeltasToListView"
'@Folder "MVVM.ValueConverters"
Option Explicit

Public Sub Initialize(ByVal ListView As MSComctlLib.ListView, ByVal Member As TtDeltaType)
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

Private Sub SetColumnHeaders(ByVal ListView As MSComctlLib.ListView, ByVal Member As TtDeltaType)
    With ListView
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="#", Width:=ListView.Width - 16
    End With
    
    If Member = ttDelta Then
        ListView.ColumnHeaders.Item(1).Text = "Before"
        ListView.ColumnHeaders.Add Text:="After"
        ListView.ColumnHeaders.Item(1).Width = (ListView.Width - 16) / 2
        ListView.ColumnHeaders.Item(2).Width = (ListView.Width - 16) / 2
    End If
End Sub

Public Sub Load(ByVal ListView As MSComctlLib.ListView, ByVal ViewModel As DeltasPreviewViewModel, _
    ByVal Member As TtDeltaType)
    ListView.ListItems.Clear
    
    'If TransferDeltasViewModel Is Nothing Then Exit Sub
    Debug.Assert Not ViewModel Is Nothing
    Debug.Assert Not ViewModel.Deltas Is Nothing
    
    Dim SourceNumberFormats As Object
    Set SourceNumberFormats = ViewModel.GetNumberFormats(ttSource)
    
    Dim DestinationNumberFormats As Object
    Set DestinationNumberFormats = ViewModel.GetNumberFormats(ttDestination)
    
    Dim Item As Variant
    If Member = ttDelta Then
        Dim TransferDelta As TransferDelta
        For Each TransferDelta In ViewModel.Deltas
            AddItemTransferDelta ListView, TransferDelta, SourceNumberFormats(TransferDelta.FieldSource), DestinationNumberFormats(TransferDelta.FieldDestination)
        Next TransferDelta
    ElseIf Member = ttKeyMember Then
        AddSelectAll ListView
        For Each Item In ViewModel.Keys
            AddItem ListView, Item
        Next Item
        UpdateHeader ListView, Member, (ViewModel.Keys.Count - 1)
    ElseIf Member = ttField Then
        AddSelectAll ListView
        For Each Item In ViewModel.Fields
            AddItem ListView, Item
        Next Item
        UpdateHeader ListView, Member, (ViewModel.Fields.Count - 1)
    End If
End Sub

Private Sub AddItemTransferDelta(ByVal ListView As MSComctlLib.ListView, ByVal TransferDelta As TransferDelta, _
    ByVal SourceNumberFormat As String, ByVal DestinationNumberFormat As String)
    Dim ValueBeforeFormatted As String
    Dim ValueAfterFormatted As String
    
    Select Case TransferDelta.DataType
        Case vbError
            ValueBeforeFormatted = CStr(TransferDelta.ValueBefore)
            ValueAfterFormatted = ERR_CAPTION
        Case vbDouble
            ValueBeforeFormatted = Format$(TransferDelta.ValueBefore, DestinationNumberFormat)
            ValueAfterFormatted = Format$(TransferDelta.ValueAfter, SourceNumberFormat)
        Case vbString
            ValueBeforeFormatted = TransferDelta.ValueBefore
            ValueAfterFormatted = TransferDelta.ValueAfter
        Case Else
            Err.Raise StringConstants.ERR_NUM_UNEXPECTED_VARTYPE, ERR_SOURCE, ERR_MSG_UNEXPECTED_VARTYPE
    End Select
    
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=ValueBeforeFormatted)
    ListItem.ListSubItems.Add Text:=ValueAfterFormatted
End Sub

Private Sub AddItem(ByVal ListView As MSComctlLib.ListView, ByVal Text As String)
    ListView.ListItems.Add Text:=Text
End Sub

Private Sub UpdateHeader(ByVal ListView As MSComctlLib.ListView, ByVal Member As TtDeltaType, ByVal Count As Long)
    Dim HeaderText As String
    
    Select Case Member
        Case ttKeyMember
            HeaderText = KEY_HEADER
        Case ttField
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
    ListItem.ForeColor = StringConstants.COLOR_SELECT_ALL
    ListItem.Key = SELECT_ALL
End Sub


