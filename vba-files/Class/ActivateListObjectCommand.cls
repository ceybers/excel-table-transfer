VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActivateListObjectCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TableProps.Commands"
Option Explicit

Private Type TState
    ListObject As ListObject
    CanExecute As Boolean
End Type
Private This As TState

Private Property Get Workbook() As Workbook
    Set Workbook = This.ListObject.Parent.Parent
End Property

Private Property Get Worksheet() As Worksheet
    Set Worksheet = This.ListObject.Parent
End Property

Public Function CanExecute() As Boolean
    If ActiveSheet.Name <> Worksheet.Name Then CanExecute = True
    If ActiveWorkbook.Name <> Workbook.Name Then CanExecute = True
End Function

Public Sub Execute()
    Workbook.Activate
    Worksheet.Activate
    This.ListObject.DataBodyRange.Cells(1, 1).Select
End Sub

Public Sub Create(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
End Sub

Public Sub UpdateCommandButton(ByVal CommandButton As CommandButton)
    If CanExecute Then
        CommandButton.Caption = "Activate"
        CommandButton.Enabled = True
    Else
        CommandButton.Caption = "Activated"
        CommandButton.Enabled = False
    End If
End Sub
