VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnQualityViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ColumnQuality"
Option Explicit
Implements IViewModel

Private Type TClsColumnQualityViewModel
    ListColumn As ListColumn
    Results As KeyColumn
End Type

Private this As TClsColumnQualityViewModel

' Property
Public Property Get ListColumn() As ListColumn
    Set ListColumn = this.ListColumn
End Property

Public Property Set ListColumn(ByVal RHS As ListColumn)
    Set this.ListColumn = RHS
    Results = KeyColumn.FromColumn(RHS)
End Property

Public Property Get Results() As KeyColumn
    Results = this.Results
End Property

Public Property Let Results(ByVal RHS As KeyColumn)
    Set this.Results = RHS
End Property

' Functions
Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = Not ListColumn Is Nothing
End Function

Public Sub UpdateListView(ByVal lv As ListView)
    Debug.Assert IViewModel_IsValid
    
    Dim target As Long
    target = this.Results.Range.Cells.Count
    
    AddListViewItem lv, "Distinct", this.Results.Count, target
    AddListViewItem lv, "Unique", this.Results.UniqueKeys.Count, target
    AddListViewItem lv, "Non-Text", this.Results.NonTextCount, 0
    AddListViewItem lv, "Errors", this.Results.ErrorCount, 0
    AddListViewItem lv, "Blanks", this.Results.BlankCount, 0
    AddListViewItem lv, "Count", this.Results.Range.Cells.Count, -1
    
    With lv.ListItems(lv.ListItems.Count)
        .Bold = True
        .ListSubItems(1).Bold = True
    End With
End Sub

Private Sub AddListViewItem(ByVal lv As ListView, ByVal caption As String, ByVal value As Integer, ByVal target As Long)
    Dim icon As String
    
    If value = target Then
        icon = "Tick"
    ElseIf value < target And target > 0 Then
        icon = "TraceError"
    Else
        icon = "Cross"
    End If
    
    If target = -1 Then
        icon = "AutoSum"
    End If
    
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
