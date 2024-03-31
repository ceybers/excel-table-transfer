VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableTableNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model2.AvailableTables")
Option Explicit

Public NodeType As TransferTreeviewNode
Public Key As String
Public ParentKey As String
Public Caption As String
Public IsSelected As Boolean
Public IsHidden As Boolean
Public IsProtected As Boolean
Public Object As Object
