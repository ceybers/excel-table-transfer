VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.KeyMapper.Model"
Option Explicit

Private Type TKeyColumn
    ListColumn As ListColumn
    Properties As ColumnProperty
    IsPreferred As Boolean
    
    ' TODO Consider impl IListable?
    Selected As Boolean
    Checked As Boolean
    Enabled As Boolean
End Type
Private This As TKeyColumn

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
End Property

Public Property Get DataType() As String
    DataType = This.Properties.ColumnType
End Property

Public Property Get IsText() As Boolean
    IsText = (GetVarTypeOfColumnRange(This.ListColumn.DataBodyRange) = vbString)
End Property

Public Property Get IsUnique() As Boolean
    IsUnique = ListColumnPropHelpers.ColumnIsUnique(This.ListColumn)
End Property

Public Property Get IsPreferred() As Boolean
    IsPreferred = This.IsPreferred
End Property

Public Property Let IsPreferred(ByVal vNewValue As Boolean)
    This.IsPreferred = vNewValue
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
Public Sub Load(ByVal ListColumn As ListColumn)
    Debug.Assert Not Me Is KeyColumn
    Set This.ListColumn = ListColumn
    Set This.Properties = New ColumnProperty
    This.Properties.Load ListColumn
End Sub

Public Function Create(ByVal ListColumn As ListColumn) As KeyColumn
    Dim Result As KeyColumn
    Set Result = New KeyColumn
    With Result
        .Load ListColumn
    End With
    Set Create = Result
End Function