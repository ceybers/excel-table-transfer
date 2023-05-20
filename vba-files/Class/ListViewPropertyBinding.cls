VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListViewPropertyBinding"
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
Private WithEvents mListView As MSComctlLib.ListView
Attribute mListView.VB_VarHelpID = -1

Private Type TState
    BindingMode As BindingMode
    ViewModel As Object
    SourcePropertyPath As String
    TargetProperty As String
    OneTimeBoundCounter As Long
    ValueConverter As IValueConverter
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
    Set IPropertyBinding_Target = mListView
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = This.TargetProperty
End Property

Public Sub Create(ByVal Source As Object, ByVal SourcePropertyPath As String, ByVal Target As Control, _
    Optional ByVal TargetProperty As String, Optional ByVal BindingMode As BindingMode, _
    Optional ByVal ValueConverter As IValueConverter)
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    
    Set This.ViewModel = Source
    This.SourcePropertyPath = SourcePropertyPath
    Set mListView = Target
    This.TargetProperty = TargetProperty
    
    This.BindingMode = BindingMode
    
    If Not ValueConverter Is Nothing Then
        Set This.ValueConverter = ValueConverter
    End If
    
    TryUpdateTarget
End Sub

'Private Sub mListView_DblClick()
'    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding mListView_DblClick()"
'
'    TryUpdateSource
'End Sub

Private Sub mListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding mListView_ItemCheck()"
    
    TryUpdateSource
End Sub

Private Sub mListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding mListView_ItemClick()"
    
    TryUpdateSource
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding IHandlePropertyChanged_HandlePropertyChanged()"
    If Not Source Is This.ViewModel Then Exit Sub
    If PropertyName <> This.SourcePropertyPath Then Exit Sub
    
    TryUpdateTarget
End Sub

Private Sub TryUpdateTarget()
    If Not CanUpdateTarget Then Exit Sub
    
    If This.TargetProperty = "SelectedItem" Then
        UpdateListViewSelectedItem
    ElseIf This.TargetProperty = "ListItems" Then
        UpdateListViewListItems
    Else
        Debug.Print "Not yet implemented"
        Debug.Assert False
    End If
    
    If This.BindingMode = OneTimeBinding Then This.OneTimeBoundCounter = This.OneTimeBoundCounter + 1
End Sub

Private Function CanUpdateTarget() As Boolean
    If This.BindingMode = TwoWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneTimeBinding And This.OneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

Private Sub TryUpdateSource()
    If Not CanUpdateSource Then Exit Sub
    
    If This.TargetProperty = "SelectedItem" Then
        UpdateSource
    Else
        UpdateSource
        'Debug.Assert False
    End If
End Sub

Private Function CanUpdateSource() As Boolean
    If This.BindingMode = TwoWayBinding Then CanUpdateSource = True
    If This.BindingMode = OneWayToSource Then CanUpdateSource = True
End Function

Private Sub UpdateListViewSelectedItem()
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding UpdateListViewSelectedItem()"
    
    Dim vNewValue As Variant
    vNewValue = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
    
    Dim ListItem As ListItem
    For Each ListItem In mListView.ListItems
        If ListItem.Text = CStr(vNewValue) Then
            'Set mListView.SelectedItem = ListItem
            mListView.SelectedItem = ListItem.Key
            Exit Sub
        End If
    Next ListItem
End Sub

Private Sub UpdateListViewListItems()
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding UpdateListViewListItems()"
      
    Dim NewList As Object
    Set NewList = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
        
    If Not This.ValueConverter Is Nothing Then
        This.ValueConverter.Convert Source:=NewList, Target:=mListView
    Else
        ' TODO Move to Default Value Converter
        mListView.ListItems.Clear
        Dim Key As Variant
        For Each Key In NewList.Keys
            mListView.ListItems.Add Key:=Key, Text:=NewList.Item(Key)
        Next Key
    End If
End Sub

Private Sub UpdateSource()
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding UpdateSource()"
    
    Dim vNewValue As Variant
    Set vNewValue = CallByName(mListView, This.TargetProperty, VbGet)
    
    If Not This.ValueConverter Is Nothing Then
        This.ValueConverter.ConvertBack Me
    Else
        CallByName This.ViewModel, This.SourcePropertyPath, VbLet, vNewValue
    End If
End Sub
