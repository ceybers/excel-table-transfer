VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboBoxPropertyBinding"
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
Private WithEvents mComboBox As MSForms.ComboBox
Attribute mComboBox.VB_VarHelpID = -1

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
    Set IPropertyBinding_Target = mComboBox
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = This.TargetProperty
End Property

Public Sub Create(ByVal Source As Object, ByVal SourcePropertyPath As String, ByVal Target As Control, _
    Optional ByVal TargetProperty As String, Optional ByVal BindingMode As BindingMode, _
    Optional ByVal ValueConverter As IValueConverter)
    If DO_DEBUG Then Debug.Print "ComboBoxPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    
    Set This.ViewModel = Source
    This.SourcePropertyPath = SourcePropertyPath
    
    Set mComboBox = Target
    This.TargetProperty = TargetProperty
    
    This.BindingMode = BindingMode
    
    UpdateTarget
End Sub

Private Sub mComboBox_Change()
    If DO_DEBUG Then Debug.Print "ComboBoxPropertyBinding cbo_Change()"
    UpdateSource
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If PropertyName <> This.SourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "ComboBoxPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If CanUpdateTarget Then
        If This.TargetProperty = "Value" Then
            UpdateDefault
        ElseIf This.TargetProperty = "List" Then
            UpdateComboBoxList
        Else
            UpdateDefault
        End If
    End If
    
    If This.BindingMode = OneTimeBinding Then This.OneTimeBoundCounter = This.OneTimeBoundCounter + 1
End Sub

Private Sub UpdateSource()
    If This.BindingMode = TwoWayBinding Or This.BindingMode = OneWayToSource Then
        Dim vNewValue As Variant
        vNewValue = CallByName(mComboBox, This.TargetProperty, VbGet)
        CallByName This.ViewModel, This.SourcePropertyPath, VbLet, vNewValue
    End If
End Sub

Private Function CanUpdateTarget() As Boolean
    If This.BindingMode = TwoWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneTimeBinding And This.OneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

Private Sub UpdateDefault()
    Dim vNewValue As Variant
    vNewValue = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
    CallByName mComboBox, This.TargetProperty, VbLet, vNewValue
End Sub

Private Sub UpdateComboBoxList()
    If DO_DEBUG Then Debug.Print "ComboBoxPropertyBinding UpdateComboBoxList()"
    
    mComboBox.Clear
    
    Dim NewList As Scripting.Dictionary
    Set NewList = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
    
    Dim Key As Variant
    For Each Key In NewList.Keys
        mComboBox.AddItem NewList.Item(Key)
    Next Key
End Sub

