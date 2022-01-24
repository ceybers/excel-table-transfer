VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColumnPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("HelperFunctions")
Option Explicit

Private Type TClsColumnPair
    lhs As ListColumn
    RHS As ListColumn
End Type

Private this As TClsColumnPair

Public Property Get lhs() As ListColumn
    Set lhs = this.lhs
End Property

Public Property Set lhs(ByVal RHS As ListColumn)
    Set this.lhs = RHS
End Property

Public Property Get RHS() As ListColumn
    Set RHS = this.RHS
End Property

Public Property Set RHS(ByVal RHS As ListColumn)
    Set this.RHS = RHS
End Property

Public Function IsValid() As Boolean
    If this.lhs Is Nothing Then Exit Function
    If this.RHS Is Nothing Then Exit Function
    IsValid = True
End Function

Public Function Create(ByRef lhs As ListColumn, ByRef RHS As ListColumn) As clsColumnPair
    Set Create = New clsColumnPair
    Set Create.lhs = lhs
    Set Create.RHS = RHS
End Function

Public Function Equals(ByRef RHS As clsColumnPair)
    If RHS Is Nothing Then Exit Function
    If Not this.lhs Is RHS.lhs Then Exit Function
    If Not this.RHS Is RHS.RHS Then Exit Function
    Equals = True
End Function

Public Function ToString() As String
    Debug.Assert Not lhs Is Nothing
    Debug.Assert Not RHS Is Nothing
    ToString = "<" & lhs.Name & ", " & RHS.Name & ">"
End Function
