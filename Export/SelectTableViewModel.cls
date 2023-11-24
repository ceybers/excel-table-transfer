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

Private Const ICON_SIZE As Integer = 16

Public Event CollectionChanged()
Public Event ItemSelected()

Private Type TViewModel
    Tables As Collection
    SelectedTable As ListObject
    ActiveTable As ListObject
    SearchCriteria As String
    
    AutoSelected As Boolean
End Type

Private This As TViewModel

Public Property Get CanSelect() As Boolean
    If This.Tables Is Nothing Then
        CanSelect = False
        Exit Property
    End If
    
    If This.Tables.Count < 2 Then
        CanSelect = False
        Exit Property
    End If
    
    CanSelect = True
End Property

Public Property Get AutoSelected() As Boolean
    AutoSelected = This.AutoSelected
End Property

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
    
    ' Check if there are exactly 2 tables and one is already set as active
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

Private Sub TrySelectOther()
    If This.ActiveTable Is Nothing Then
        Exit Sub
    End If
    
    If This.Tables.Count <> 2 Then
        Exit Sub
    End If
    
    Dim lo As ListObject
    For Each lo In This.Tables
        If lo.Range.Address(external:=True) <> This.ActiveTable.Range.Address(external:=True) Then
            Set SelectedTable = lo
            This.AutoSelected = True
            Exit Sub
        End If
    Next lo
End Sub

Private Sub class_initialize()
    This.AutoSelected = False
    Load
End Sub

Private Function IViewModel_IsValid() As Boolean
    If This.SelectedTable Is Nothing Then Exit Function
    IViewModel_IsValid = True
End Function

Public Sub LoadTreeview(ByVal tv As MSComctlLib.TreeView)
    With tv
        .ImageList = GetMSOImageList(ICON_SIZE)
        
        .Nodes.Clear
        .Nodes.Add key:="Root", text:="Excel", image:="Excel"
        
        .LineStyle = tvwTreeLines
        .Appearance = cc3D
        .Style = tvwTreelinesPlusMinusPictureText
        .Indentation = ICON_SIZE
        .LabelEdit = tvwManual
        .HideSelection = False
    End With
    
    Dim lo As ListObject
    For Each lo In This.Tables
        TryAddNode tv, lo.parent.parent
        TryAddNode tv, lo.parent
        TryAddNode tv, lo
    Next lo
    
    Dim nd As Node
    For Each nd In tv.Nodes
        nd.Expanded = True
    Next nd
End Sub

Private Sub TryAddNode(ByVal tv As MSComctlLib.TreeView, ByVal obj As Object)
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim wb As Workbook
    
    Dim key As String
    Dim parent As String
    Dim nd As Node
    Dim image As String
    Dim text As String
    Dim suffix As String
    
    If TypeOf obj Is Workbook Then
        Set wb = obj
        key = "[" & wb.Name & "]"
        parent = "Root"
        image = "wb"
        text = wb.Name
        
    ElseIf TypeOf obj Is Worksheet Then
        Set ws = obj
        key = "[" & ws.parent.Name & "]" & ws.Name
        parent = "[" & ws.parent.Name & "]"
        image = "ws"
        text = ws.Name
        
    ElseIf TypeOf obj Is ListObject Then
        Set lo = obj
        key = lo.Range.Address(external:=True)
        parent = "[" & lo.parent.parent.Name & "]" & lo.parent.Name
        image = "lo"
        text = lo.Name
        suffix = vbNullString
        
        If Not This.ActiveTable Is Nothing Then
            If This.ActiveTable.Range.Address(external:=True) = lo.Range.Address(external:=True) Then
                suffix = " (active)"
                image = "activeLo"
            End If
        End If
        
        text = text & suffix
    End If
    
    For Each nd In tv.Nodes
        If nd.key = key Then Exit Sub
    Next nd
    
    tv.Nodes.Add relative:=parent, Relationship:=tvwChild, key:=key, text:=text, image:=image
    
    If image = "activeLo" Then
        tv.Nodes(key).ForeColor = RGB(192, 192, 192)
    End If
End Sub

