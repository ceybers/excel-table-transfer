VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TState
    ListColumn As ListColumn
    Selected As Boolean
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.ListColumn.Name
End Property

Public Property Get Selected() As Boolean
    Selected = This.Selected
End Property

Public Property Let Selected(ByVal vNewValue As Boolean)
    This.Selected = vNewValue
End Property

Public Property Get ColumnType() As String
    ColumnType = VarTypeValueToString(GetVarTypeOfColumnRange(This.ListColumn.DataBodyRange))
End Property

Public Property Get R1C1() As String
    R1C1 = GetR1C1(This.ListColumn)
End Property

Public Property Get Blanks() As Long
    Blanks = ColumnHasBlanks(This.ListColumn)
End Property

Public Property Get Errors() As Long
    Errors = ColumnHasErrors(This.ListColumn)
End Property

Public Property Get Locked() As Long
    Locked = ColumnIsLocked(This.ListColumn)
End Property

Public Property Get Formula() As Long
    Formula = ColumnHasFormulae(This.ListColumn)
End Property

Public Property Get Validation() As Long
    Validation = ColumnHasValidation(This.ListColumn)
End Property

Public Property Get Unique() As Long
    Unique = ColumnIsUnique(This.ListColumn)
End Property
    
Public Sub Load(ByVal ListColumn As ListColumn)
    Set This.ListColumn = ListColumn
End Sub

Public Sub Activate()
    This.ListColumn.DataBodyRange.Select
End Sub

Private Function VarTypeValueToString(ByVal Value As Long) As String
    Select Case Value
        Case vbEmpty:
            VarTypeValueToString = "Empty"
        Case vbDouble:
            VarTypeValueToString = "Number"
        Case vbString:
            VarTypeValueToString = "Text"
        Case vbDate:
            VarTypeValueToString = "Date"
        Case vbCurrency:
            VarTypeValueToString = "Currency"
        Case vbVariant:
            VarTypeValueToString = "Mixed"
    End Select
End Function
