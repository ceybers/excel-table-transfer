VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandlePropertyChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder MVVM.Infrastructure.Abstract
Option Explicit

'@Description "A callback invoked when a property value is set."
Public Sub HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
Attribute HandlePropertyChanged.VB_Description = "A callback invoked when a property value is set."
End Sub
