VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferDeltasToListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.ValueConverters")
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    If TypeOf Source Is TransferDeltas Then
        Load Source, Target
    Else
        LoadCollection Source, Target
    End If
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding)
End Sub

Private Sub LoadCollection(ByVal Collection As Collection, ByVal ListView As MScomctllib.ListView)
    ListView.ListItems.Clear
    
    AddSelectAll ListView
    
    Dim Item As Variant
    For Each Item In Collection
        AddItem ListView, Item
    Next Item
End Sub

Private Sub Load(ByVal Deltas As TransferDeltas, ByVal ListView As MScomctllib.ListView)
    ListView.ListItems.Clear
    
    Dim SourceNumberFormats As Object
    Set SourceNumberFormats = Deltas.SourceNumberFormats
    
    Dim DestinationNumberFormats As Object
    Set DestinationNumberFormats = Deltas.DestinationNumberFormats
    
    Dim TransferDelta As TransferDelta
    For Each TransferDelta In Deltas.Items
        AddItemTransferDelta ListView, TransferDelta, SourceNumberFormats(TransferDelta.FieldSource), DestinationNumberFormats(TransferDelta.FieldDestination)
    Next TransferDelta
End Sub

Private Sub AddItemTransferDelta(ByVal ListView As MScomctllib.ListView, ByVal TransferDelta As TransferDelta, _
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

Private Sub AddItem(ByVal ListView As MScomctllib.ListView, ByVal Text As String)
    ListView.ListItems.Add Text:=Text
End Sub

Private Sub UpdateHeader(ByVal ListView As MScomctllib.ListView, ByVal Member As TtDeltaType, ByVal Count As Long)
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

Private Sub AddSelectAll(ByVal ListView As MScomctllib.ListView)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=SELECT_ALL)
    ListItem.ForeColor = StringConstants.COLOR_SELECT_ALL
    ListItem.Key = SELECT_ALL
End Sub

Private Sub IValueConverter_Initialize(ByVal Target As Object)
    Dim ListView As MScomctllib.ListView
    Set ListView = Target
    With ListView
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
        .HideSelection = False
    End With
    
    'SetColumnHeaders ListView, Member
End Sub

Private Sub SetColumnHeaders(ByVal ListView As MScomctllib.ListView, ByVal Member As TtDeltaType)
    With ListView
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="#", Width:=ListView.Width - 16
    End With
    
    If Member = ttDelta Then
        ListView.ColumnHeaders.Item(1).Text = COL_HDR_BEFORE
        ListView.ColumnHeaders.Add Text:=COL_HDR_AFTER
        ListView.ColumnHeaders.Item(1).Width = (ListView.Width - 16) / 2
        ListView.ColumnHeaders.Item(2).Width = (ListView.Width - 16) / 2
    End If
End Sub
