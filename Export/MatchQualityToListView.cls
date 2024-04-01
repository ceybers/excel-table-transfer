VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchQualityToListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.ValueConverters")
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object) ' read
    Load Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding) ' write
    ' Read-only
End Sub

Public Sub Load(ByVal KeyColumn As Collection, ByVal ListView As MScomctllib.ListView)
    'Debug.Assert Not Collection Is Nothing
    
    ListView.ListItems.Clear
    
    Dim Item As Variant
    For Each Item In KeyColumn
        ListView.ListItems.Add Text:=CStr(Item)
    Next Item
End Sub

Public Sub Initialize(ByVal ListView As MScomctllib.ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
    End With
End Sub



