VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PickableTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.TablePicker.Model")
Option Explicit

Private Type TPickableTables
    Item As Collection ' of type PickableTable
    ShowSmallTables As Boolean
    ShowProtectedWorksheets As Boolean
End Type
Private This As TPickableTables

Private Sub Class_Initialize()
    Set This.Item = New Collection
End Sub

'@Enumerator
Public Property Get Item() As Collection
Attribute Item.VB_UserMemId = -4
    Set Item = This.Item
End Property

Public Sub Load(ByVal Application As Application)
    Dim Workbook As Workbook
    Dim Worksheet As Worksheet
    Dim ListObject As ListObject
    
    For Each Workbook In Application.Workbooks
        For Each Worksheet In Workbook.Worksheets
            For Each ListObject In Worksheet.ListObjects
                This.Item.Add PickableTable.Create(ListObject)
            Next ListObject
        Next Worksheet
    Next Workbook
End Sub

Private Function IsFiltered(ByVal PickableTable As PickableTable) As Boolean
    ' check if we are hiding protected files and this file is protected
    ' etc etc
End Function

Public Function GetByKey(ByVal Key As String) As PickableTable
    Dim PickableTable As PickableTable
    For Each PickableTable In This.Item
        If PickableTable.Key = Key Then
            Set GetByKey = PickableTable
            Exit Function
        End If
    Next PickableTable
End Function
