Attribute VB_Name = "TestKeyMapper"
'@Folder "Tests.MVVM"
Option Explicit
Option Private Module

Public Sub TestKeyMapper()
    Dim ti As TransferInstruction
    Set ti = GetTestTransferInstruction
    
    Dim vm As KeyMapperViewModel
    Set vm = New KeyMapperViewModel
    'vm.LoadFromTransferInstruction ti
    
    Dim View As IView
    Set View = New KeyMapperView
    
    'Set vm.LHSTable = Nothing ' Fails because property guards against null
    'Set vm.RHSTable = Nothing ' Fails because property guards against null
    
    Set vm.LHSTable = ti.Source
    Set vm.RHSTable = ti.Destination
    
    If View.ShowDialog(vm) Then
        Debug.Print "ShowDialog true"
    Else
        Debug.Print "ShowDialog false"
    End If
End Sub

Public Sub TestKeyMapperView()
    Debug.Assert True
    
    Dim vm As KeyMapperViewModel
    Dim View As IView
    
    Set vm = New KeyMapperViewModel
    
    Set View = New KeyMapperView
    
    ' TODO Fix
    Dim vview As KeyMapperView
    Set vview = KeyMapperView
    vview.DEBUG_EVENTS = True
    
    Set vm.LHSTable = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    Set vm.RHSTable = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
    
    If View.ShowDialog(vm) Then
        Debug.Print "ShowDialog true"
    Else
        Debug.Print "ShowDialog false"
    End If
End Sub

Public Sub TestRemoveUnmappedKeys()
    Dim comp As KeyColumnComparer
    
    Set comp = New KeyColumnComparer
    Set comp.LHS = KeyColumn.FromColumn(ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Item(1))
    Set comp.RHS = KeyColumn.FromColumn(ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2).ListColumns.Item(2))

    Dim mapResult As Variant
    mapResult = comp.Map
    
    RemoveUnmappedKeys comp, mapResult
End Sub

Public Sub TestAppendUnmappedKeys()
    Dim comp As KeyColumnComparer
    
    Set comp = New KeyColumnComparer
    Set comp.LHS = KeyColumn.FromColumn(ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Item(1))
    Set comp.RHS = KeyColumn.FromColumn(ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2).ListColumns.Item(2))

    'Dim mapResult As Variant
    'mapResult = comp.Map
    comp.Map
    
    AppendUnmappedKeys comp
End Sub

Public Sub RemoveUnmappedKeys(ByVal comp As KeyColumnComparer, Optional ByRef cachedMappedResults As Variant)
    If IsMissing(cachedMappedResults) Then
        cachedMappedResults = comp.Map
    End If
   
    Dim i As Long
    Dim rng As Range
    Set rng = comp.RHS.Range
    
    For i = rng.rows.Count To 1 Step -1
        If cachedMappedResults(rng.rows.Item(i).row) = -1 Then
            rng.rows.Item(i).EntireRow.Delete
        End If
    Next i
End Sub

Public Sub AppendUnmappedKeys(ByVal comp As KeyColumnComparer)
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim lr As ListRow
    Dim i As Long
     
    Set lc = GetListColumnFromRange(comp.RHS.Range)
    Set lo = lc.parent

    For i = 1 To comp.LeftOnly.Count
        Set lr = lo.ListRows.Add(alwaysinsert:=True)
        lr.Range.Cells.Item(1, lc.Index).Value2 = comp.LeftOnly.Item(i)
    Next i
End Sub

