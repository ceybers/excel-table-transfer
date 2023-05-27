VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICancellable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Describes an object that encapsulates 'IsCancelled' state."
'@Interface
'@Folder("MVVM.Infrastructure.Abstract")
'@ModuleDescription "Describes an object that encapsulates 'IsCancelled' state."
Option Explicit

'@Description "Gets a value indicating whether the object is in a cancelled state."
Public Property Get IsCancelled() As Boolean
Attribute IsCancelled.VB_Description = "Gets a value indicating whether the object is in a cancelled state."
End Property

'@Description "Sets the IsCancelled flag to True."
Public Sub OnCancel()
Attribute OnCancel.VB_Description = "Sets the IsCancelled flag to True."
End Sub
