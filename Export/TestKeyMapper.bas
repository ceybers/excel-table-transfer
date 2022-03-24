Attribute VB_Name = "TestKeyMapper"
'@Folder("KeyMapper")
Option Explicit
Option Private Module

Public Sub Test()
    Dim vm As KeyMapperViewModel
    Dim view As IView
    
    Set vm = New KeyMapperViewModel
    
    Set view = New KeyMapperView
    
    ' TODO Fix
    Dim vview As KeyMapperView
    Set vview = KeyMapperView
    vview.DEBUG_EVENTS = True
    
    Set vm.LHSTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set vm.RHSTable = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    If view.ShowDialog(vm) Then
        Debug.Print "ShowDialog true"
    Else
        Debug.Print "ShowDialog false"
    End If
End Sub

Public Sub TestRemoveUnmappedKeys()
    Dim comp As KeyColumnComparer
    
    Set comp = New KeyColumnComparer
    Set comp.lhs = KeyColumn.FromColumn(ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1))
    Set comp.rhs = KeyColumn.FromColumn(ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(2))

    Dim mapResult As Variant
    mapResult = comp.Map
    
    RemoveUnmappedKeys comp, mapResult
End Sub

Public Sub TestAppendUnmappedKeys()
    Dim comp As KeyColumnComparer
    
    Set comp = New KeyColumnComparer
    Set comp.lhs = KeyColumn.FromColumn(ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1))
    Set comp.rhs = KeyColumn.FromColumn(ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(2))

    Dim mapResult As Variant
    mapResult = comp.Map
    
    AppendUnmappedKeys comp
End Sub

Public Sub RemoveUnmappedKeys(ByVal comp As KeyColumnComparer, Optional ByRef cachedMappedResults As Variant)
    If IsMissing(cachedMappedResults) Then
        cachedMappedResults = comp.Map
    End If
   
    Dim i As Long
    Dim rng As Range
    Set rng = comp.rhs.Range
    
    For i = rng.rows.Count To 1 Step -1
        If cachedMappedResults(rng.rows(i).row) = -1 Then
             rng.rows(i).EntireRow.Delete
        End If
    Next i
End Sub

Public Sub AppendUnmappedKeys(ByVal comp As KeyColumnComparer)
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim lr As ListRow
    Dim i As Long
     
    Set lc = GetListColumnFromRange(comp.rhs.Range)
    Set lo = lc.parent

    For i = 1 To comp.LeftOnly.Count
        Set lr = lo.ListRows.Add(alwaysinsert:=True)
        lr.Range.Cells(1, lc.Index).Value2 = comp.LeftOnly.Item(i)
    Next i
End Sub

