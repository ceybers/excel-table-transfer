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
    rhs As ListColumn
End Type

Private this As TClsColumnPair

Public Property Get LHS() As ListColumn
    Set LHS = this.LHS
End Property

Public Property Set LHS(ByVal rhs As ListColumn)
    Set this.LHS = rhs
End Property

Public Property Get rhs() As ListColumn
    Set rhs = this.rhs
End Property

Public Property Set rhs(ByVal rhs As ListColumn)
    Set this.rhs = rhs
End Property

Public Function IsValid() As Boolean
    If this.LHS Is Nothing Then Exit Function
    If this.rhs Is Nothing Then Exit Function
    IsValid = True
End Function

Public Function Create(ByVal LHS As ListColumn, ByVal rhs As ListColumn) As ColumnPair
    Set Create = New ColumnPair
    With Create
        Set .LHS = LHS
        Set .rhs = rhs
    End With
End Function

Public Function Equals(ByVal rhs As ColumnPair) As Boolean
    If rhs Is Nothing Then Exit Function
    If Not this.LHS Is rhs.LHS Then Exit Function
    If Not this.rhs Is rhs.rhs Then Exit Function
    Equals = True
End Function

Public Function ToString() As String
    Debug.Assert Not LHS Is Nothing
    Debug.Assert Not rhs Is Nothing
    ToString = LHS.Name & "," & rhs.Name
End Function
