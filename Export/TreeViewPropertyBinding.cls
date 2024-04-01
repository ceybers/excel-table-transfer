VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeViewPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder "MVVM.Infrastructure.Bindings.PropertyBindings"
Option Explicit
Implements IPropertyBinding
Implements IHandlePropertyChanged

Private Const DO_DEBUG As Boolean = False

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mTreeView As MScomctllib.TreeView
Attribute mTreeView.VB_VarHelpID = -1

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
    Set IPropertyBinding_Target = mTreeView
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = This.TargetProperty
End Property

Public Sub Create(ByVal Source As Object, ByVal SourcePropertyPath As String, ByVal Target As Control, _
    Optional ByVal TargetProperty As String, Optional ByVal BindingMode As BindingMode, _
    Optional ByVal ValueConverter As IValueConverter)
    If DO_DEBUG Then Debug.Print "TreeViewPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    
    Set This.ViewModel = Source
    This.SourcePropertyPath = SourcePropertyPath
    
    Set mTreeView = Target
    This.TargetProperty = TargetProperty
    
    If TargetProperty = "SelectedItem" Then
        This.BindingMode = TwoWayBinding
    Else
        This.BindingMode = OneWayBinding
    End If
    
    UpdateTarget
End Sub

Private Sub mTreeView_NodeClick(ByVal Node As MScomctllib.Node)
    If DO_DEBUG Then Debug.Print "TreeViewPropertyBinding tv_NodeClick()"
    UpdateSource
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If PropertyName <> This.SourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "TreeViewPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If CanUpdateTarget Then
        If This.TargetProperty = "SelectedItem" Then
            UpdateTreeViewSelectedItem
        ElseIf This.TargetProperty = "Nodes" Then
            UpdateTreeViewNodes
        Else
            Debug.Print "Not yet implemented"
            Debug.Assert False
        End If
    End If
    
    If This.BindingMode = OneTimeBinding Then This.OneTimeBoundCounter = This.OneTimeBoundCounter + 1
End Sub

Private Sub UpdateSource()
    If This.BindingMode = TwoWayBinding Or This.BindingMode = OneWayToSource Then
        If This.TargetProperty = "SelectedItem" Then
            Dim SelectedNode As Node
            Set SelectedNode = CallByName(mTreeView, This.TargetProperty, VbGet)
            Debug.Assert Not SelectedNode Is Nothing
            CallByName This.ViewModel, This.SourcePropertyPath, VbLet, SelectedNode.Text
        Else
            Debug.Assert False
        End If
    End If
End Sub

Private Function CanUpdateTarget() As Boolean
    If This.BindingMode = TwoWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneTimeBinding And This.OneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

Private Sub UpdateTreeViewSelectedItem()
    Dim vNewValue As String
    vNewValue = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
    
    Dim Node As Node
    For Each Node In mTreeView.Nodes
        If Node.Text = vNewValue Then
            Set mTreeView.SelectedItem = Node
            Exit Sub
        End If
    Next Node
End Sub

Private Sub UpdateTreeViewNodes()
    If DO_DEBUG Then Debug.Print "TreeViewPropertyBinding UpdateTreeViewNodes()"
    
    mTreeView.Nodes.Clear
    
    Dim NewList As Scripting.Dictionary
    Set NewList = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
    
    Dim Key As Variant
    For Each Key In NewList.Keys
        mTreeView.Nodes.Add Key:=Key, Text:=NewList.Item(Key)
    Next Key
End Sub

