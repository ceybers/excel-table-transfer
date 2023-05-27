VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Abstract
Option Explicit

Public Function CreateBinding(ByVal Source As Object, _
    ByVal SourcePropertyPath As String, _
    ByVal Target As Control, _
    Optional ByVal TargetProperty As String, _
    Optional ByVal BindingMode As BindingMode, _
    Optional ByVal ValueConverter As IValueConverter) _
    As IPropertyBinding
    
End Function
