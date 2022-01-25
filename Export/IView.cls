VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Abstractions"
'@ModuleDescription "Describes a View."
'@Interface
'@Exposed
Option Explicit

'@Description "Displays the View modally and returns a value indicating confirmation (True) or cancellation (False)."
Public Function ShowDialog(ByVal ViewModel As IViewModel) As Boolean

End Function
