VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPair2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model2.ColumnPair2"
Option Explicit

Public Source As String
Public Destination As String

Public Function Exists(ByVal RHS As ColumnPair2) As Boolean
    Exists = (Source = RHS.Source And Destination = RHS.Destination)
End Function
