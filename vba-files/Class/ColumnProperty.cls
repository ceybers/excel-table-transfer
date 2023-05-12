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
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.ListColumn.Name
End Property

Public Property Get ColumnType() As String
    ColumnType = "Mixed"
End Property

Public Property Get R1C1() As String
    R1C1 = GetR1C1(This.ListColumn)
End Property

Public Property Get Blanks() As String
    Blanks = ColumnHasBlanks(This.ListColumn)
End Property

Public Property Get Errors() As String
    Errors = ColumnHasErrors(This.ListColumn)
End Property

Public Property Get Locked() As String
    Locked = ColumnIsLocked(This.ListColumn)
End Property

Public Property Get Formula() As String
    Formula = ColumnHasFormulae(This.ListColumn)
End Property

Public Property Get Validation() As String
    Validation = ColumnHasValidation(This.ListColumn)
End Property

Public Property Get Unique() As String
    Unique = "Mixed"
End Property
    
Public Sub Load(ByVal ListColumn As ListColumn)
    Set This.ListColumn = ListColumn
End Sub
