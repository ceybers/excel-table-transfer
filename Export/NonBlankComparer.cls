VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonBlankComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.ValueComparer.Strategies"
Option Explicit
Implements IValueComparer

Private Function IValueComparer_Compare(ByVal LHS As Variant, ByVal RHS As Variant) As Boolean
    If VarType(RHS) = vbError Then Exit Function
    
    Select Case VarType(LHS)
        Case vbError:
            Exit Function
        Case vbEmpty:
            Exit Function
        Case vbString:
            If Len(LHS) = 0 Then Exit Function
    End Select
    
    IValueComparer_Compare = True
End Function
