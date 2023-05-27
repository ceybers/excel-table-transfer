VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotifyPropertyChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ParameterNotUsed
'@Folder MVVM.Infrastructure.Abstract
Option Explicit

'@Description "Registers the specified handler."
Public Sub RegisterHandler(ByVal Handler As IHandlePropertyChanged)
Attribute RegisterHandler.VB_Description = "Registers the specified handler."
End Sub

'@Description "Notifies all registered handlers of a property value change."
Public Sub OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
Attribute OnPropertyChanged.VB_Description = "Notifies all registered handlers of a property value change."
End Sub
