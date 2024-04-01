VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionButtonBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Infrastructure.Bindings.PropertyBindings.Strategies"
Option Explicit
Implements IBindingStrategy

Private Function IBindingStrategy_CreateBinding(ByVal Source As Object, ByVal SourcePropertyPath As String, _
    ByVal Target As Control, Optional ByVal TargetProperty As String, Optional ByVal BindingMode As BindingMode, _
    Optional ByVal ValueConverter As IValueConverter) As IPropertyBinding
    Dim NewBinding As OptionButtonPropertyBinding
    Set NewBinding = New OptionButtonPropertyBinding
    NewBinding.Create Source, SourcePropertyPath, Target, TargetProperty, BindingMode, ValueConverter
    Set IBindingStrategy_CreateBinding = NewBinding
End Function
