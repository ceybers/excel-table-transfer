VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Abstract
Option Explicit

Public Enum BindingMode
    TwoWayBinding
    OneWayBinding
    OneWayToSource
    OneTimeBinding
End Enum

Public Property Get Mode() As BindingMode
End Property

'@Description "The ViewModel's Property that is bound to the UserForm's Control."
Public Property Get Source() As IViewModel
Attribute Source.VB_Description = "The ViewModel's Property that is bound to the UserForm's Control."
End Property

Public Property Get SourcePropertyPath() As String
End Property

'@Description "The UserForm Control that is bound to the ViewModel's Property."
Public Property Get Target() As Object ' UserForm Control
Attribute Target.VB_Description = "The UserForm Control that is bound to the ViewModel's Property."
End Property

Public Property Get TargetProperty() As String
End Property
