VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ColumnPairs"
Option Explicit

Private Type TClsColumnPair
    lhs As ListColumn
    rhs As ListColumn
End Type

Private This As TClsColumnPair

Public Property Get lhs() As ListColumn
    Set lhs = This.lhs
End Property

Public Property Set lhs(ByVal rhs As ListColumn)
    Set This.lhs = rhs
End Property

Public Property Get rhs() As ListColumn
    Set rhs = This.rhs
End Property

Public Property Set rhs(ByVal rhs As ListColumn)
    Set This.rhs = rhs
End Property

Public Function IsValid() As Boolean
    If This.lhs Is Nothing Then Exit Function
    If This.rhs Is Nothing Then Exit Function
    IsValid = True
End Function

Public Function Create(ByVal lhs As ListColumn, ByVal rhs As ListColumn) As ColumnPair
    Set Create = New ColumnPair
    With Create
        Set .lhs = lhs
        Set .rhs = rhs
    End With
End Function

Public Function Equals(ByVal rhs As ColumnPair) As Boolean
    If rhs Is Nothing Then Exit Function
    If Not This.lhs Is rhs.lhs Then Exit Function
    If Not This.rhs Is rhs.rhs Then Exit Function
    Equals = True
End Function

Public Function ToString() As String
    Debug.Assert Not lhs Is Nothing
    Debug.Assert Not rhs Is Nothing
    ToString = lhs.Name & "," & rhs.Name
End Function
