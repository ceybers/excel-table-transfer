VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Bindings
Option Explicit
Implements IHandlePropertyChanged

Private Type TBindingManager
    Context As IAppContext
    PropertyBindings As Collection
    BindingStrategies As Scripting.Dictionary
    HandledViewModels As Scripting.Dictionary
End Type
Private This As TBindingManager

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Friend Property Set Context(ByVal vNewValue As IAppContext)
    Set This.Context = vNewValue
End Property

Private Sub Class_Initialize()
    Set This.PropertyBindings = New Collection
    
    Set This.BindingStrategies = New Scripting.Dictionary
    InitializeBindingStrategies
    
    Set This.HandledViewModels = New Scripting.Dictionary
End Sub

Private Sub InitializeBindingStrategies()
    With This.BindingStrategies
        .Add Item:=New TextBoxBindingStrategy, Key:="MSForms.TextBox"
        .Add Item:=New LabelBindingStrategy, Key:="MSForms.Label"
        .Add Item:=New CheckBoxBindingStrategy, Key:="MSForms.CheckBox"
        .Add Item:=New OptionButtonBindingStrategy, Key:="MSForms.OptionButton"
        .Add Item:=New ComboBoxBindingStrategy, Key:="MSForms.ComboBox"
        .Add Item:=New ListViewBindingStrategy, Key:="MSComctlLib.ListView"
        .Add Item:=New TreeViewBindingStrategy, Key:="MSComctlLib.TreeView"
        .Add Item:=New CommandButtonBindingStrategy, Key:="MSForms.CommandButton"
        .Add Item:=New FrameBindingStrategy, Key:="MSForms.Frame"
    End With
End Sub

Public Sub BindPropertyPath(ByVal Source As Object, ByVal SourcePropertyPath As String, ByVal Target As Object, _
    Optional ByVal TargetProperty As String, Optional ByVal BindingMode As BindingMode, _
    Optional ByVal ValueConverter As IValueConverter)
    Dim BindingStrategy As IBindingStrategy
    
    On Error GoTo NoStrategyFound
    Set BindingStrategy = This.BindingStrategies.Item(GetFullTypeName(Target))
    On Error GoTo 0
    
    Dim ThisBindingPath As IBindingPath
    Set ThisBindingPath = BindingPath.Create(Source, SourcePropertyPath)
    
    Dim NewBinding As IPropertyBinding
    Set NewBinding = BindingStrategy.CreateBinding(ThisBindingPath.Object, ThisBindingPath.PropertyName, _
        Target, TargetProperty, BindingMode, ValueConverter)
    This.PropertyBindings.Add Item:=NewBinding
    
    Dim SourceNotifier As INotifyPropertyChanged
    Set SourceNotifier = ThisBindingPath.Object
    If Not SourceNotifier Is Nothing And TypeOf NewBinding Is IHandlePropertyChanged Then
        '@Ignore ArgumentWithIncompatibleObjectType
        SourceNotifier.RegisterHandler NewBinding
    End If
    
    If TypeOf ThisBindingPath.Object Is INotifyPropertyChanged Then
        'Dim SourceTypeName As String
        'SourceTypeName = TypeName(ThisBindingPath.Object)
        Dim SourceInstance As Object
        Set SourceInstance = ThisBindingPath.Object
        If Not This.HandledViewModels.Exists(SourceInstance) Then
            Dim Notifier As INotifyPropertyChanged
            Set Notifier = ThisBindingPath.Object
            Notifier.RegisterHandler Me
            This.HandledViewModels.Add Key:=SourceInstance, Item:=SourceInstance
        End If
    End If
    
    Exit Sub
NoStrategyFound:
    Err.Clear
    MsgBox "This control hasn't been implemented yet." & vbCrLf & GetFullTypeName(Target)
End Sub

Private Function GetFullTypeName(ByVal Target As Control) As String
    GetFullTypeName = TypeName(Target)
   
    Select Case GetFullTypeName
        Case "ListView4":
            GetFullTypeName = "MSComctlLib.ListView"
        Case "TreeView4":
            GetFullTypeName = "MSComctlLib.TreeView"
        Case Else:
            If TypeOf Target Is MSForms.Control Then
                GetFullTypeName = "MSForms." & GetFullTypeName
            End If
    End Select
End Function

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    'This.Context.CommandManager.EvaluateCanExecute Source
End Sub
