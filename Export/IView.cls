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
'@Interface
Option Explicit

' TODO implement from https://github.com/rubberduck-vba/MVVM/blob/main/src/IView.cls
Public Function ShowDialog(ByVal viewModel As IViewModel) As Boolean
End Function
