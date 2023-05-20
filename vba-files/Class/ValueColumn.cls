VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.ValueMapper.Model"
Option Explicit

Private Type TValueColumn
    ListColumn As ListColumn
    Properties As ColumnProperty
    
    MappedTo As String
    
    IsKey As Boolean
    IsStarred As Boolean
    
    ' TODO Consider impl IListable?
    Selected As Boolean
    Checked As Boolean
    Enabled As Boolean
End Type
Private This As TValueColumn

Public Property Get ListColumn() As ListColumn
    Set ListColumn = This.ListColumn
End Property

Public Property Get Key() As String
    Key = This.ListColumn.Name
End Property

'@Description "Name of the underlying ListObject."
Public Property Get Name() As String
Attribute Name.VB_Description = "Name of the underlying ListObject."
    Name = This.ListColumn.Name
    If This.IsKey Then
        Name = Name & " (key column)"
    End If
End Property

Public Property Get MappedTo() As String
    MappedTo = This.MappedTo
End Property

Public Property Let MappedTo(ByVal vNewValue As String)
    This.MappedTo = vNewValue
End Property

Public Property Get DataType() As String
    DataType = This.Properties.ColumnType
End Property

Public Property Get IsHidden() As Boolean
    IsHidden = (This.ListColumn.DataBodyRange.ColumnWidth = 0)
End Property

Public Property Get IsText() As Boolean
    IsText = (GetVarTypeOfColumnRange(This.ListColumn.DataBodyRange) = vbString)
End Property

Public Property Get IsUnique() As Boolean
    IsUnique = ListColumnPropHelpers.ColumnIsUnique(This.ListColumn)
End Property

Public Property Get IsProtected() As Boolean
    IsProtected = (ListColumnPropHelpers.ColumnIsLocked(This.ListColumn) <> none)
End Property
' ---
Public Property Get IsKey() As Boolean
    IsKey = This.IsKey
End Property
 
Public Property Let IsKey(ByVal vNewValue As Boolean)
    This.IsKey = vNewValue
End Property

Public Property Get IsStarred() As Boolean
    IsStarred = This.IsStarred
End Property
 
Public Property Let IsStarred(ByVal vNewValue As Boolean)
    This.IsStarred = vNewValue
End Property

' ---
Public Property Get Selected() As Boolean
    Selected = This.Selected
End Property

Public Property Let Selected(ByVal vNewValue As Boolean)
    This.Selected = vNewValue
End Property

Public Property Get Checked() As Boolean
    Checked = This.Checked
End Property

Public Property Let Checked(ByVal vNewValue As Boolean)
    This.Checked = vNewValue
End Property

Public Property Get Enabled() As Boolean
    Enabled = This.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    This.Enabled = vNewValue
End Property

' ---
Public Function Create(ByVal ListColumn As ListColumn) As ValueColumn
    Dim Result As ValueColumn
    Set Result = New ValueColumn
    With Result
        .Load ListColumn
    End With
    Set Create = Result
End Function

Public Sub Load(ByVal ListColumn As ListColumn)
    Debug.Assert Not Me Is ValueColumn
    Set This.ListColumn = ListColumn
    Set This.Properties = New ColumnProperty
    This.Properties.Load ListColumn
    
    AssessEnabled
End Sub

Private Sub AssessEnabled()
    This.Enabled = True
    
    If This.Properties.ColumnType <> "Text" And This.Properties.ColumnType <> "Empty" Then
        This.Enabled = False
        'Debug.Assert False
    Else
        'Debug.Assert False
    End If
End Sub
