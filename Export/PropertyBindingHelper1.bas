Attribute VB_Name = "PropertyBindingHelper1"
'@Folder("Helpers")
Option Explicit

Public Sub DebugPrintPropertyBinding(ByVal PropertyBinding As IPropertyBinding)
    Debug.Print "PropertyBinding, Mode: "; PropertyBinding.Mode
    Debug.Print " Source: "; TypeName(PropertyBinding.Source); ", Path: "; PropertyBinding.SourcePropertyPath
    Debug.Print " Target: "; TypeName(PropertyBinding.Target); ", Path: "; PropertyBinding.TargetProperty; vbNullString
End Sub
