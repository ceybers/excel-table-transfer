Attribute VB_Name = "modMain"
Option Explicit

Public Sub TableTransfer()
    If Selection.ListObject Is Nothing Then Exit Sub
    
    With New clsTableTransfer
        .TransferFrom Selection.ListObject
    End With
End Sub


