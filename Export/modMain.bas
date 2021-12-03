Attribute VB_Name = "modMain"
Option Explicit

Public Sub TableTransfer()
    Dim transfer As clsTableTransfer
    
    If Selection.ListObject Is Nothing Then Exit Sub
    
    Set transfer = New clsTableTransfer
    transfer.TransferFrom Selection.ListObject
End Sub


