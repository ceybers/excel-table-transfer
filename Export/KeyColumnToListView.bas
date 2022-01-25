Attribute VB_Name = "KeyColumnToListView"
'@Folder("KeyColumn")
Option Explicit


Public Sub UpdateListView(ByVal lv As ListView)
    Debug.Assert IViewModel_IsValid
    
    AddListViewItem lv, "Distinct", this.Results.Count, "Tick"
    AddListViewItem lv, "Unique", this.Results.UniqueKeys.Count, "Tick"
    AddListViewItem lv, "Non-Text", this.Results.NonTextCount, "Cross"
    AddListViewItem lv, "Errors", this.Results.ErrorCount, "TraceError"
    AddListViewItem lv, "Blanks", this.Results.BlankCount, "Cross"
    AddListViewItem lv, "Count", this.Results.Range.Cells.Count, "AutoSum"
    
    With lv.ListItems(lv.ListItems.Count)
        .Bold = True
        .ListSubItems(1).Bold = True
    End With
End Sub

Private Sub AddListViewItem(ByVal lv As ListView, ByVal caption As String, ByVal value As Integer, ByVal icon As String)
    With lv.ListItems.Add(text:=caption, icon:=icon, SmallIcon:=icon)
        .ListSubItems.Add text:=value
    End With
End Sub

Public Sub InitializeListView(ByVal lv As ListView4)
    SetListViewImageList lv
    SetListViewProperties lv
End Sub

Private Sub SetListViewProperties(ByVal lv As ListView)
    With lv
        .view = lvwReport
        .HideSelection = False
        .CheckBoxes = False
        .LabelEdit = lvwManual
        .Gridlines = True
        .BorderStyle = ccNone
    End With
    
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    
    lv.ColumnHeaders.Add text:="Description"
    lv.ColumnHeaders.Add text:="Value"
    lv.ColumnHeaders(2).Alignment = lvwColumnRight
    lv.ColumnHeaders(2).Width = (72 / 2)
    lv.ColumnHeaders(1).Width = lv.Width - (72 / 2) - 5
End Sub

Private Sub SetListViewImageList(ByVal lv As ListView)
    Dim il As ImageList
    'If lv.Icons Is Nothing Then
    If True Then
        Set il = GetMSOImageList
        Set lv.Icons = GetMSOImageList(32)
        Set lv.SmallIcons = GetMSOImageList(16)
    End If
End Sub
