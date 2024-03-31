VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnQualityViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ViewModels"
Option Explicit
Implements IViewModel

Private Const MSO_PASS As String = "lblComplete"
Private Const MSO_WARNING As String = "lblWarning"
Private Const MSO_ERROR As String = "lblError"
Private Const MSO_TOTAL As String = "lblAutoSum"

Private Type TClsColumnQualityViewModel
    ListColumn As ListColumn
    Results As KeyColumn
End Type

Private This As TClsColumnQualityViewModel

' Property
Public Property Get ListColumn() As ListColumn
    Set ListColumn = This.ListColumn
End Property

Public Property Set ListColumn(ByVal RHS As ListColumn)
    Set This.ListColumn = RHS
    'Results = KeyColumn.FromColumn(RHS)
    Set This.Results = KeyColumn.FromColumn(RHS)
End Property

Public Property Get Results() As KeyColumn
    Set Results = This.Results
End Property

Public Property Let Results(ByVal RHS As KeyColumn)
    Set This.Results = RHS
End Property

' Functions
Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = Not ListColumn Is Nothing
End Function

Public Sub UpdateListView(ByVal lv As ListView)
    Debug.Assert IViewModel_IsValid
    
    If This.Results Is Nothing Then
        AddListViewItem lv, "Empty Table", 0, -2
        Exit Sub
    End If
    
    Dim Target As Long
    Target = This.Results.Range.Cells.Count
    
    AddListViewItem lv, "Distinct", This.Results.Count, Target
    AddListViewItem lv, "Unique", This.Results.UniqueKeys.Count, Target
    AddListViewItem lv, "Non-Text", This.Results.NonTextCount, 0
    AddListViewItem lv, "Errors", This.Results.ErrorCount, 0
    AddListViewItem lv, "Blanks", This.Results.BlankCount, 0
    AddListViewItem lv, "Count", This.Results.Range.Cells.Count, -1
    
    With lv.ListItems.Item(lv.ListItems.Count)
        .Bold = True
        .ListSubItems.Item(1).Bold = True
    End With
End Sub

Private Sub AddListViewItem(ByVal lv As ListView, ByVal Caption As String, ByVal Value As Long, ByVal Target As Long)
    Dim icon As String
    
    If Value = Target Then
        icon = MSO_PASS
    ElseIf Value < Target And Target > 0 Then
        icon = MSO_WARNING
    Else
        icon = MSO_ERROR
    End If
    
    If Target = -1 Then
        icon = MSO_TOTAL
    End If
    
    With lv.ListItems.Add(Text:=Caption, icon:=icon, SmallIcon:=icon)
       .ListSubItems.Add Text:=Value
    End With
End Sub

Public Sub InitializeListView(ByVal lv As ListView4)
    SetListViewImageList lv
    SetListViewProperties lv
End Sub

Private Sub SetListViewProperties(ByVal lv As ListView)
    With lv
        .View = lvwReport
        .HideSelection = False
        .CheckBoxes = False
        .LabelEdit = lvwManual
        .Gridlines = True
        .BorderStyle = ccNone
    End With
    
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    
    lv.ColumnHeaders.Add Text:="Description"
    lv.ColumnHeaders.Add Text:="Value"
    lv.ColumnHeaders.Item(2).Alignment = lvwColumnRight
    lv.ColumnHeaders.Item(2).Width = (72 / 2)
    lv.ColumnHeaders.Item(1).Width = lv.Width - (72 / 2) - 5
End Sub

Private Sub SetListViewImageList(ByVal lv As ListView)
    'Dim il As ImageList
    'If lv.Icons Is Nothing Then
    'If True Then
        'Set il = GetMSOImageList
        Set lv.Icons = GetMSOImageList(16) 'GetMSOImageList(32)
        Set lv.SmallIcons = GetMSOImageList(16)
    'End If
End Sub

