VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColumnQualityViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ColumnQuality"
Option Explicit
Implements IViewModel

Private Type TClsColumnQualityViewModel
    ListColumn As ListColumn
    Results As ArrayExAnalyseOne
End Type

Private this As TClsColumnQualityViewModel

' Property
Public Property Get ListColumn() As ListColumn
    Set ListColumn = this.ListColumn
End Property

Public Property Set ListColumn(ByVal RHS As ListColumn)
    Set this.ListColumn = RHS
    
    Dim arr As Variant
    arr = RHS.DataBodyRange
    Results = ArrayAnalyseOne(arr)
End Property

Public Property Get Results() As ArrayExAnalyseOne
    Results = this.Results
End Property

Public Property Let Results(ByRef RHS As ArrayExAnalyseOne)
    this.Results = RHS
End Property

' Functions
Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = Not ListColumn Is Nothing
End Function

Public Sub UpdateListView(ByRef lv As ListView)
    Debug.Assert IViewModel_IsValid
    
    AddListViewItem lv, "Distinct", this.Results.Distinct, "AcceptInvitation"
    AddListViewItem lv, "Unique", this.Results.Unique, "AcceptInvitation"
    AddListViewItem lv, "Errors", this.Results.Errors, "TraceError"
    AddListViewItem lv, "Blanks", this.Results.Blanks, "DeclineInvitation"
    AddListViewItem lv, "Count", this.Results.length, "AutoSum"
    
    With lv.ListItems(lv.ListItems.Count)
        .Bold = True
        .ListSubItems(1).Bold = True
    End With
End Sub

Private Sub AddListViewItem(ByRef lv As ListView, ByVal caption As String, ByVal value As Integer, ByVal icon As String)
    With lv.ListItems.Add(text:=caption, icon:=icon, SmallIcon:=icon)
        .ListSubItems.Add text:=value
    End With
End Sub

Public Sub InitializeListView(ByRef lv As ListView)
    SetListViewImageList lv
    SetListViewProperties lv
End Sub

Private Sub SetListViewProperties(ByRef lv As ListView)
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

Private Sub SetListViewImageList(ByRef lv As ListView)
    Dim il As ImageList
    If lv.Icons Is Nothing Then
        Set il = GetMSOImageList
        Set lv.Icons = GetMSOImageList(32)
        Set lv.SmallIcons = GetMSOImageList(16)
    End If
End Sub
