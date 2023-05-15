VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
Option Explicit

Public Sub Load(Optional ByVal Object As Object)
End Sub

Public Sub Save()
End Sub

Public Sub Reset()
End Sub

Public Function GetFlag(ByVal FlagName As String) As Boolean
End Function

Public Sub SetFlag(ByVal FlagName As String, ByVal FlagValue As Boolean)
End Sub

