VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectTableViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ViewModels"
Option Explicit
Option Compare Text
Implements IViewModel

Private Const ICON_SIZE As Long = 16

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

Public Property Set SelectedTable(ByVal RHS As ListObject)
    Set This.SelectedTable = RHS
    RaiseEvent ItemSelected
End Property

Public Property Get ActiveTable() As ListObject
    Set ActiveTable = This.ActiveTable
End Property

Public Property Set ActiveTable(ByVal RHS As ListObject)
    Set This.ActiveTable = RHS
    Load
End Property

Public Property Get SearchCriteria() As String
    SearchCriteria = This.SearchCriteria
End Property

Public Property Let SearchCriteria(ByVal RHS As String)
    This.SearchCriteria = RHS
    Load
End Property

Private Sub Load()
    Set This.Tables = GetAllTablesInApplication
    
    ' Check if there are exactly 2 tables and one is already set as active
    TrySelectOther
    
    RaiseEvent CollectionChanged
End Sub

Public Sub TrySelect(ByVal Key As String)
    Dim lo As ListObject
    
    If Not This.ActiveTable Is Nothing Then
        If This.ActiveTable.Range.Address(External:=True) = Key Then
            ' Set selected table to nothing, otherwise we can still double click the active table and
            ' it will close the dialog and return the previously selected valid table
            Set SelectedTable = Nothing
            Exit Sub
        End If
    End If
    
    For Each lo In This.Tables
        If lo.Range.Address(External:=True) = Key Then
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
        If lo.Range.Address(External:=True) <> This.ActiveTable.Range.Address(External:=True) Then
            Set SelectedTable = lo
            This.AutoSelected = True
            Exit Sub
        End If
    Next lo
End Sub

Private Sub Class_Initialize()
    This.AutoSelected = False
    Load
End Sub

Private Function IViewModel_IsValid() As Boolean
    If This.SelectedTable Is Nothing Then Exit Function
    IViewModel_IsValid = True
End Function

Public Sub LoadTreeview(ByVal tv As MSComctlLib.TreeView)
    With tv
        Set .ImageList = GetMSOImageList(ICON_SIZE)
        
        .Nodes.Clear
        .Nodes.Add Key:="Root", Text:="Excel", image:="Excel"
        
        .LineStyle = tvwTreeLines
        .Appearance = cc3D
        .Style = tvwTreelinesPlusMinusPictureText
        .Indentation = ICON_SIZE
        .LabelEdit = tvwManual
        .HideSelection = False
    End With
    
    Dim lo As ListObject
    For Each lo In This.Tables
        TryAddNode tv, lo.Parent.Parent
        TryAddNode tv, lo.Parent
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
    
    Dim Key As String
    Dim Parent As String
    Dim nd As Node
    Dim image As String
    Dim Text As String
    Dim suffix As String
    
    If TypeOf obj Is Workbook Then
        Set wb = obj
        Key = "[" & wb.Name & "]"
        Parent = "Root"
        image = "wb"
        Text = wb.Name
        
    ElseIf TypeOf obj Is Worksheet Then
        Set ws = obj
        Key = "[" & ws.Parent.Name & "]" & ws.Name
        Parent = "[" & ws.Parent.Name & "]"
        image = "ws"
        Text = ws.Name
        
    ElseIf TypeOf obj Is ListObject Then
        Set lo = obj
        Key = lo.Range.Address(External:=True)
        Parent = "[" & lo.Parent.Parent.Name & "]" & lo.Parent.Name
        image = "lo"
        Text = lo.Name
        suffix = vbNullString
        
        If Not This.ActiveTable Is Nothing Then
            If This.ActiveTable.Range.Address(External:=True) = lo.Range.Address(External:=True) Then
                suffix = " (active)"
                image = "activeLo"
            End If
        End If
        
        Text = Text & suffix
    End If
    
    For Each nd In tv.Nodes
        If nd.Key = Key Then Exit Sub
    Next nd
    
    tv.Nodes.Add Relative:=Parent, relationship:=tvwChild, Key:=Key, Text:=Text, image:=image
    
    If image = "activeLo" Then
        tv.Nodes.ControlDefault(Key).ForeColor = RGB(192, 192, 192)
    End If
End Sub

