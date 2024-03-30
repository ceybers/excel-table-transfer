VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NullAlwaysComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model2.ValueComparer.Strategies"
Option Explicit
Implements IValueComparer

Private Function IValueComparer_Compare(ByVal LHS As Variant, ByVal RHS As Variant) As Boolean
    IValueComparer_Compare = True
End Function

