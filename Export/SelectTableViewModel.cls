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
    AutoSelected As Boolean
End Type

Private this As TViewModel

Public Property Get AutoSelected() As Boolean
    AutoSelected = this.AutoSelected
End Property

Public Property Get Tables() As Collection
    Set Tables = this.Tables
End Property

Public Property Get SelectedTable() As ListObject
    Set SelectedTable = this.SelectedTable
End Property

Public Property Set SelectedTable(ByVal RHS As ListObject)
    Set this.SelectedTable = RHS
    RaiseEvent ItemSelected
End Property

Public Property Get ActiveTable() As ListObject
    Set ActiveTable = this.ActiveTable
End Property

Public Property Set ActiveTable(ByVal RHS As ListObject)
    Set this.ActiveTable = RHS
    Load
End Property

Public Property Get SearchCriteria() As String
    SearchCriteria = this.SearchCriteria
End Property

Public Property Let SearchCriteria(ByVal RHS As String)
    this.SearchCriteria = RHS
    Load
End Property

Private Sub Load()
    Set this.Tables = GetAllTables
    
    TrySelectOther
    
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
                If (this.SearchCriteria = vbNullString) Or (lo.Name Like this.SearchCriteria) Then
                    GetAllTables.Add lo, lo.Range.Address(external:=True)
                End If
            Next lo
        Next ws
    Next wb
End Function

Public Sub TrySelect(ByVal key As String)
    Dim lo As ListObject
    
    If Not this.ActiveTable Is Nothing Then
        If this.ActiveTable.Range.Address(external:=True) = key Then
            ' Set selected table to nothing, otherwise we can still double click the active table and
            ' it will close the dialog and return the previously selected valid table
            Set SelectedTable = Nothing
            Exit Sub
        End If
    End If
    
    For Each lo In this.Tables
        If lo.Range.Address(external:=True) = key Then
            Set SelectedTable = lo
            Exit Sub
        End If
    Next lo
    Set SelectedTable = Nothing
End Sub

Private Sub TrySelectOther()
    If this.ActiveTable Is Nothing Then
        Exit Sub
    End If
    
    If this.Tables.Count <> 2 Then
        Exit Sub
    End If
    
    Dim lo As ListObject
    For Each lo In this.Tables
        If lo.Range.Address(external:=True) <> this.ActiveTable.Range.Address(external:=True) Then
            Set SelectedTable = lo
            this.AutoSelected = True
            Exit Sub
        End If
    Next lo
End Sub

Private Sub Class_Initialize()
    Load
End Sub

Private Function IViewModel_IsValid() As Boolean
    If this.SelectedTable Is Nothing Then Exit Function
    IViewModel_IsValid = True
End Function
