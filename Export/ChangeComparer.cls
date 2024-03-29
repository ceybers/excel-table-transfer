VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChangeComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model.TransferInstruction2"
Option Explicit
Implements IValueComparer

Private Function IValueComparer_Compare(ByVal LHS As Variant, ByVal RHS As Variant) As Boolean
    If GetVarTypeMod(LHS) <> 8 Then Exit Function
    
    If GetVarTypeMod(RHS) = 8 Then
        If LHS = RHS Then Exit Function
    End If
    
    IValueComparer_Compare = True
End Function

