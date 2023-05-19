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
    DisableProtectedWorksheets As Boolean
    DisableReadonlyWorksheets As Boolean
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
' ---
Public Property Get DisableProtectedWorksheets() As Boolean
    DisableProtectedWorksheets = This.DisableProtectedWorksheets
End Property

Public Property Let DisableProtectedWorksheets(ByVal vNewValue As Boolean)
    This.DisableProtectedWorksheets = vNewValue
End Property

Public Property Get DisableReadonlyWorksheets() As Boolean
    DisableReadonlyWorksheets = This.DisableReadonlyWorksheets
End Property

Public Property Let DisableReadonlyWorksheets(ByVal vNewValue As Boolean)
    This.DisableReadonlyWorksheets = vNewValue
End Property

' ---
Public Sub Load(ByVal Application As Application)
    Debug.Assert Not Me Is PickableTables
    
    Reload
End Sub

Public Sub Reload()
    This.Item.RemoveAll
    
    Dim Workbook As Workbook
    Dim Worksheet As Worksheet
    Dim ListObject As ListObject
    
    For Each Workbook In Application.Workbooks
        For Each Worksheet In Workbook.Worksheets
            For Each ListObject In Worksheet.ListObjects
                TryAdd PickableTable.Create(ListObject)
            Next ListObject
        Next Worksheet
    Next Workbook
End Sub

Private Sub TryAdd(ByVal PickableTable As PickableTable)
    If This.DisableProtectedWorksheets And PickableTable.WorksheetLocked Then PickableTable.Enabled = False
    If This.DisableReadonlyWorksheets And PickableTable.WorkbookReadonly Then PickableTable.Enabled = False
    
    This.Item.Add Key:=PickableTable.Key, Item:=PickableTable
End Sub

Private Function IsFiltered(ByVal PickableTable As PickableTable) As Boolean
    ' check if we are hiding protected files and this file is protected
    ' etc etc
    IsFiltered = False
End Function

Public Function GetByKey(ByVal Key As String) As PickableTable
    Set GetByKey = This.Item.Item(Key)
End Function

Public Function GetByListObject(ByVal ListObject As ListObject) As PickableTable
    Dim TempPickableTable As PickableTable
    Set TempPickableTable = PickableTable.Create(ListObject)
    
    If Not This.Item.Exists(TempPickableTable.Key) Then Exit Function
    Set GetByListObject = This.Item.Item(TempPickableTable.Key)
End Function
