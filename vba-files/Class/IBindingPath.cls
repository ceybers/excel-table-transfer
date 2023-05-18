VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder MVVM.Infrastructure.Abstract
Option Explicit

'@Description "Gets the context for resolving the Path."
Public Property Get Context() As Object
Attribute Context.VB_Description = "Gets the context for resolving the Path."
End Property

'@Description "Gets the property path of the binding."
Public Property Get Path() As String
Attribute Path.VB_Description = "Gets the property path of the binding."
End Property

'@Description "Gets a reference to the binding source object resolved from the Path."
Public Property Get Object() As Object
Attribute Object.VB_Description = "Gets a reference to the binding source object resolved from the Path."
End Property

'@Description "Gets the name of the bound property resolved from the Path."
Public Property Get PropertyName() As String
Attribute PropertyName.VB_Description = "Gets the name of the bound property resolved from the Path."
End Property
