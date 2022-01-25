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
    RHS As ListColumn
End Type

Private This As TClsColumnPair

Public Property Get lhs() As ListColumn
    Set lhs = This.lhs
End Property

Public Property Set lhs(ByVal RHS As ListColumn)
    Set This.lhs = RHS
End Property

Public Property Get RHS() As ListColumn
    Set RHS = This.RHS
End Property

Public Property Set RHS(ByVal RHS As ListColumn)
    Set This.RHS = RHS
End Property

Public Function IsValid() As Boolean
    If This.lhs Is Nothing Then Exit Function
    If This.RHS Is Nothing Then Exit Function
    IsValid = True
End Function

Public Function Create(ByVal lhs As ListColumn, ByVal RHS As ListColumn) As ColumnPair
    Set Create = New ColumnPair
    With Create
        Set .lhs = lhs
        Set .RHS = RHS
    End With
End Function

Public Function Equals(ByVal RHS As ColumnPair) As Boolean
    If RHS Is Nothing Then Exit Function
    If Not This.lhs Is RHS.lhs Then Exit Function
    If Not This.RHS Is RHS.RHS Then Exit Function
    Equals = True
End Function

Public Function ToString() As String
    Debug.Assert Not lhs Is Nothing
    Debug.Assert Not RHS Is Nothing
    ToString = "<" & lhs.Name & ", " & RHS.Name & ">"
End Function
