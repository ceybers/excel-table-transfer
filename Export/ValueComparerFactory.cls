VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueComparerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Model.ValueComparer.Factory"
Option Explicit

' VarType 8 (and 5 and 11) only
Public Property Get Default()
    Set Default = New ChangeComparer
End Property

' Ignore errors, Empty, and Len=0 strings
Public Property Get NonBlank()
    Set Default = New NonBlankComparer
End Property

' Always transfer everything
Public Property Get NullAlways()
    'Set NullAlways = New NullAlwaysComparer
End Property

' Ignore errors on LHS or RHS, always transfer everything else
Public Property Get Overwrite()
    'Set Overwrite = New OverwriteComparer
End Property
