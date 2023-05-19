Attribute VB_Name = "TestWorkflow"
'@IgnoreModule ImplicitActiveWorkbookReference
'@Folder "TableTransfer"
Option Explicit

Private ctx As IAppContext
Private SrcTable As ListObject
Private DstTable As ListObject
Private SrcKeyColumn As ListColumn
Private DstKeyColumn As ListColumn
Private MappedValueColumns As Collection

Public Sub DoTestWorkflow()
    Debug.Print "START"
    Debug.Print "---"
    
    Worksheets(1).Activate
    Range("A2").Activate
    
    Set ctx = New AppContext
    Set SrcTable = Nothing
    Set DstTable = Nothing
    
    If PickDirection(Selection.ListObject) Then
        Debug.Print "PickDirection OK"
    Else
        Debug.Print "PickDirection EXIT"
        Exit Sub
    End If
    
    If PickOtherTable() Then
        Debug.Print "PickOtherTable OK"
    Else
        Debug.Print "PickOtherTable EXIT"
        Exit Sub
    End If
    
    If PickKeys() Then
        Debug.Print "PickKeys OK"
    Else
        Debug.Print "PickKeys EXIT"
        Exit Sub
    End If
    
    If PickValues() Then
        Debug.Print "PickValues OK"
    Else
        Debug.Print "PickValues EXIT"
        Exit Sub
    End If
    
    Debug.Print "END"
    Debug.Print "---"
End Sub

Private Function PickDirection(ByVal ListObject As ListObject) As Boolean
    Dim DirectionViewModel As DirectionPickerViewModel
    Set DirectionViewModel = New DirectionPickerViewModel
    DirectionViewModel.Load ListObject

    Dim DirectionView As IView
    Set DirectionView = DirectionPickerView.Create(ctx, DirectionViewModel)
    
    Dim Result As Boolean
    Result = DirectionView.ShowDialog
    
    If Result Then
        If DirectionViewModel.Result = Source Then
            Set SrcTable = ListObject
        ElseIf DirectionViewModel.Result = Destination Then
            Set DstTable = ListObject
        Else
            Debug.Assert False
        End If
    End If
    
    Set DirectionViewModel = Nothing
    Set DirectionView = Nothing
    PickDirection = Result
End Function

Private Function PickOtherTable() As Boolean
    Dim TableMapperVM As TableMapperViewModel
    Set TableMapperVM = New TableMapperViewModel
    TableMapperVM.Load SrcTable, DstTable
    
    Dim TableMapperV As IView
    Set TableMapperV = TableMapperView.Create(ctx, TableMapperVM)
    
    Dim Result As Boolean
    Result = TableMapperV.ShowDialog
    
    If Result Then
        Set SrcTable = TableMapperVM.SrcTableVM.Selected.ListObject
        Set DstTable = TableMapperVM.DstTableVM.Selected.ListObject
    End If
    
    Set TableMapperVM = Nothing
    Set TableMapperV = Nothing
    PickOtherTable = Result
End Function

Private Function PickKeys() As Boolean
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim KeyMapperVM As KeyMapperViewModel
    Set KeyMapperVM = New KeyMapperViewModel
    KeyMapperVM.Load SrcTable:=SrcTable, _
        DstTable:=DstTable

    Dim KeyMapperV As IView
    Set KeyMapperV = KeyMapperView.Create(ctx, KeyMapperVM)
    
    Dim Result As Boolean
    Result = KeyMapperV.ShowDialog
    
    If Result Then
        Debug.Print "   KeyMapper result:"
        Debug.Print "      Src: "; KeyMapperVM.SrcKeyColumnVM.SelectedAsText
        Debug.Print "       Dst: "; KeyMapperVM.DstKeyColumnVM.SelectedAsText
        Set SrcKeyColumn = KeyMapperVM.SrcKeyColumnVM.Selected.ListColumn
        Set DstKeyColumn = KeyMapperVM.SrcKeyColumnVM.Selected.ListColumn
    End If
    
    Set KeyMapperVM = Nothing
    Set KeyMapperV = Nothing
    PickKeys = Result
End Function

Private Function PickValues() As Boolean
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim ValueMapperVM As ValueMapperViewModel
    Set ValueMapperVM = New ValueMapperViewModel
    
    ValueMapperVM.Load _
        SrcColumn:=SrcKeyColumn, _
        DstColumn:=DstKeyColumn

    Dim ValueMapperV As IView
    Set ValueMapperV = ValueMapperView.Create(ctx, ValueMapperVM)
    
    Dim Result As Boolean
    Result = ValueMapperV.ShowDialog
    
    If Result Then
        Debug.Print "   ValueMapper result:"
        Debug.Print "      "; ValueMapperVM.MappedValueColumns.Count; " tuples"
        
        Dim ColumnTuple As ColumnTuple
        For Each ColumnTuple In ValueMapperVM.MappedValueColumns
            Debug.Print "         "; ColumnTuple.SourceListColumn.Name; " -> "; ColumnTuple.DestinationListColumn.Name
        Next ColumnTuple
        
        Set MappedValueColumns = ValueMapperVM.MappedValueColumns
    End If
    
    Set ValueMapperVM = Nothing
    Set ValueMapperV = Nothing
    PickValues = Result
End Function
