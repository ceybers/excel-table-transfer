VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OverwriteComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model.ValueComparer.Strategies"
Option Explicit
Implements IValueComparer

Private Function IValueComparer_Compare(ByVal LHS As Variant, ByVal RHS As Variant) As Boolean
    If VarType(LHS) = vbError Then Exit Function
    If VarType(RHS) = vbError Then Exit Function
    
    IValueComparer_Compare = True
End Function
