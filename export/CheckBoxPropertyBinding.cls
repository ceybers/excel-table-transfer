VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckBoxPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder MVVM.Infrastructure.Bindings.PropertyBindings
Option Explicit
Implements IPropertyBinding
Implements IHandlePropertyChanged

Private Const DO_DEBUG As Boolean = False

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mCheckBox As MSForms.CheckBox
Attribute mCheckBox.VB_VarHelpID = -1

Private Type TState
    BindingMode As BindingMode
    ViewModel As Object
    SourcePropertyPath As String
    TargetProperty As String
    OneTimeBoundCounter As Long
End Type
Private This As TState

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = This.BindingMode
End Property

Private Property Get IPropertyBinding_Source() As IViewModel
    Set IPropertyBinding_Source = This.ViewModel
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = This.SourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = mCheckBox
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = This.TargetProperty
End Property

Public Sub Create(ByVal Source As Object, ByVal SourcePropertyPath As String, ByVal Target As Control, _
    Optional ByVal TargetProperty As String, Optional ByVal BindingMode As BindingMode, _
    Optional ByVal ValueConverter As IValueConverter)
    If DO_DEBUG Then Debug.Print "CheckBoxPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    
    Set This.ViewModel = Source
    This.SourcePropertyPath = SourcePropertyPath
    
    Set mCheckBox = Target
    This.TargetProperty = TargetProperty
    
    This.BindingMode = BindingMode
    'If TargetProperty = "Value" Then
    '    This.BindingMode = TwoWayBinding
    'Else
    '    This.BindingMode = OneWayBinding
    'End If
    
    TryUpdateTarget
End Sub

Private Sub mCheckBox_Change()
    If DO_DEBUG Then Debug.Print "CheckBoxPropertyBinding cb_Change()"
    TryUpdateSource
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If PropertyName <> This.SourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "CheckBoxPropertyBinding vm_PropertyChanged()"
    TryUpdateTarget
End Sub

Private Sub TryUpdateTarget()
    If CanUpdateTarget Then
        Dim vNewValue As String
        vNewValue = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
        CallByName mCheckBox, This.TargetProperty, VbLet, vNewValue
    End If
    
    If This.BindingMode = OneTimeBinding Then This.OneTimeBoundCounter = This.OneTimeBoundCounter + 1
End Sub

Private Function CanUpdateTarget() As Boolean
    If This.BindingMode = TwoWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneTimeBinding And This.OneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

Private Sub TryUpdateSource()
    If This.BindingMode = TwoWayBinding Or This.BindingMode = OneWayToSource Then
        Dim vNewValue As Variant
        vNewValue = CallByName(mCheckBox, This.TargetProperty, VbGet)
        CallByName This.ViewModel, This.SourcePropertyPath, VbLet, vNewValue
    End If
End Sub

Private Function CanUpdateSource() As Boolean
    If This.BindingMode = TwoWayBinding Then CanUpdateSource = True
    If This.BindingMode = OneWayToSource Then CanUpdateSource = True
End Function
