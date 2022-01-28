VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectTableViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SelectTable")
Option Explicit
Option Compare Text
Implements IViewModel

Public Event CollectionChanged()
Public Event ItemSelected()

Private Type TViewModel
    SelectedTable As ListObject
    ActiveTable As ListObject
    SearchCriteria As String
    Tables As Collection
End Type

Private This As TViewModel

Public Property Get Tables() As Collection
    Set Tables = This.Tables
End Property

Public Property Get SelectedTable() As ListObject
    Set SelectedTable = This.SelectedTable
End Property

Public Property Set SelectedTable(ByVal rhs As ListObject)
    Set This.SelectedTable = rhs
    RaiseEvent ItemSelected
End Property

Public Property Get ActiveTable() As ListObject
    Set ActiveTable = This.ActiveTable
End Property

Public Property Set ActiveTable(ByVal rhs As ListObject)
    Set This.ActiveTable = rhs
    Load
End Property

Public Property Get SearchCriteria() As String
    SearchCriteria = This.SearchCriteria
End Property

Public Property Let SearchCriteria(ByVal rhs As String)
    This.SearchCriteria = rhs
    Load
End Property

Private Sub Load()
    Set This.Tables = GetAllTables
    RaiseEvent CollectionChanged
End Sub

Private Function GetAllTables() As Collection
    Set GetAllTables = New Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            For Each lo In ws.ListObjects
                If (This.SearchCriteria = vbNullString) Or (lo.Name Like This.SearchCriteria) Then
                    GetAllTables.Add lo, lo.Range.Address(external:=True)
                End If
            Next lo
        Next ws
    Next wb
End Function

Public Sub TrySelect(ByVal key As String)
    Dim lo As ListObject
    
    If Not This.ActiveTable Is Nothing Then
        If This.ActiveTable.Range.Address(external:=True) = key Then
            ' Set selected table to nothing, otherwise we can still double click the active table and
            ' it will close the dialog and return the previously selected valid table
            Set SelectedTable = Nothing
            Exit Sub
        End If
    End If
    
    For Each lo In This.Tables
        If lo.Range.Address(external:=True) = key Then
            Set SelectedTable = lo
            Exit Sub
        End If
    Next lo
    Set SelectedTable = Nothing
End Sub

Private Sub Class_Initialize()
    Load
End Sub

Private Function IViewModel_IsValid() As Boolean
    If This.SelectedTable Is Nothing Then Exit Function
    IViewModel_IsValid = True
End Function
