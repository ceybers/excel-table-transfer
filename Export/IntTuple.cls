VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntTuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public a As Integer
Public b As Integer

Function ToString() As String
    ToString = a & " -> " & b
End Function

Function PrintA()
    Debug.Print ToString()
End Function

