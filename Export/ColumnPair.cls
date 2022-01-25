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
    LHS As ListColumn
    RHS As ListColumn
End Type

Private this As TClsColumnPair

Public Property Get LHS() As ListColumn
    Set LHS = this.LHS
End Property

Public Property Set LHS(ByVal RHS As ListColumn)
    Set this.LHS = RHS
End Property

Public Property Get RHS() As ListColumn
    Set RHS = this.RHS
End Property

Public Property Set RHS(ByVal RHS As ListColumn)
    Set this.RHS = RHS
End Property

Public Function IsValid() As Boolean
    If this.LHS Is Nothing Then Exit Function
    If this.RHS Is Nothing Then Exit Function
    IsValid = True
End Function

Public Function Create(ByVal LHS As ListColumn, ByVal RHS As ListColumn) As ColumnPair
    Set Create = New ColumnPair
    With Create
        Set .LHS = LHS
        Set .RHS = RHS
    End With
End Function

Public Function Equals(ByVal RHS As ColumnPair) As Boolean
    If RHS Is Nothing Then Exit Function
    If Not this.LHS Is RHS.LHS Then Exit Function
    If Not this.RHS Is RHS.RHS Then Exit Function
    Equals = True
End Function

Public Function ToString() As String
    Debug.Assert Not LHS Is Nothing
    Debug.Assert Not RHS Is Nothing
    ToString = "<" & LHS.Name & ", " & RHS.Name & ">"
End Function
