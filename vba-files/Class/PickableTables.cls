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
'@Folder "MVVM.TableMapper.Model"
Option Explicit

Private Type TPickableTables
    Item As Scripting.Dictionary ' of type PickableTable
    ShowSmallTables As Boolean
    ShowProtectedWorksheets As Boolean
End Type
Private This As TPickableTables

Private Sub Class_Initialize()
    Set This.Item = New Scripting.Dictionary
End Sub

'@Enumerator
Public Property Get Item() As Scripting.Dictionary
Attribute Item.VB_UserMemId = -4
    Set Item = This.Item
End Property

Public Sub Load(ByVal Application As Application)
    Debug.Assert Not Me Is PickableTables
    
    Dim Workbook As Workbook
    Dim Worksheet As Worksheet
    Dim ListObject As ListObject
    
    For Each Workbook In Application.Workbooks
        For Each Worksheet In Workbook.Worksheets
            For Each ListObject In Worksheet.ListObjects
                Dim ThisPickableTable As PickableTable
                Set ThisPickableTable = PickableTable.Create(ListObject)
                This.Item.Add Key:=ThisPickableTable.Key, Item:=ThisPickableTable
            Next ListObject
        Next Worksheet
    Next Workbook
End Sub

Private Function IsFiltered(ByVal PickableTable As PickableTable) As Boolean
    ' check if we are hiding protected files and this file is protected
    ' etc etc
    IsFiltered = False
End Function

Public Function GetByKey(ByVal Key As String) As PickableTable
    Set GetByKey = This.Item.Item(Key)
End Function
