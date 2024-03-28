VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@PredeclaredId
'@Folder "Model.ColumnPairs"
Option Explicit

Private Type TClsColumnPair
    LHS As ListColumn
    RHS As ListColumn
End Type

Private This As TClsColumnPair

Public Property Get LHS() As ListColumn
    Set LHS = This.LHS
End Property

Public Property Set LHS(ByVal RHS As ListColumn)
    Set This.LHS = RHS
End Property

Public Property Get RHS() As ListColumn
    Set RHS = This.RHS
End Property

Public Property Set RHS(ByVal RHS As ListColumn)
    Set This.RHS = RHS
End Property

Public Function Create(ByVal LHS As ListColumn, ByVal RHS As ListColumn) As ColumnPair
    Dim Result As ColumnPair
    Set Result = New ColumnPair
    
    With Result
        Set .LHS = LHS
        Set .RHS = RHS
    End With
    
    Set Create = Result
End Function

Public Function IsValid() As Boolean
    If This.LHS Is Nothing Then Exit Function
    If This.RHS Is Nothing Then Exit Function
    
    IsValid = True
End Function

Public Function Equals(ByVal RHS As ColumnPair) As Boolean
    If RHS Is Nothing Then Exit Function
    If Not This.LHS Is RHS.LHS Then Exit Function
    If Not This.RHS Is RHS.RHS Then Exit Function
    
    Equals = True
End Function

Public Function ToString() As String
    Debug.Assert Not LHS Is Nothing
    Debug.Assert Not RHS Is Nothing
    
    ToString = "(" & LHS.Name & "," & RHS.Name & ")"
End Function
