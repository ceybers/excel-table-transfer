Attribute VB_Name = "AppContext"
'@IgnoreModule EmptyIfBlock
'@Folder "MVVM.AppContext"
Option Explicit

'Private Transfer As TransferInstruction
Private Transfer2 As TransferInstruction2

Private SelectTableVM As SelectTableViewModel

Private OneTableSelected As ListObject

Private GoBack As Boolean

'@ExcelHotkey e
'@EntryPoint
Public Sub TransferTable()
Attribute TransferTable.VB_ProcData.VB_Invoke_Func = "e\n14"
    PrintTime "Start", True
       
    'MsgBox "Welcome to table transfer wizard"
    
    InitializeViewModels
    
    ' DEBUG
    Set Transfer2.Source.Table = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    Set Transfer2.Destination.Table = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
    Transfer2.Source.KeyColumnName = "KeyA"
    Transfer2.Destination.KeyColumnName = "KeyB"
    GoTo Rewind3
    
    ' This won't work with TransferInstruction2
    'If TryLoadHistory Then
    '    GoTo Rewind3
    'End If
    
    ' DEBUG
    'Set Transfer.Source = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    'Set Transfer.Destination = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
    
    If CheckTablesAvailable = False Then Exit Sub
    PrintTime "CheckTablesAvailable"
    
Rewind1:
    CheckIfSelectionContainsTable
    
    If TryGetSourceOrDestination = False Then Exit Sub
    PrintTime "TryGetSourceOrDestination"
    
    If TryGetSecondTable = False Then Exit Sub
    PrintTime "TryGetSecondTable"
    
Rewind2:
    If TryGetKeyColumns = False Then Exit Sub
    If GoBack Then
        GoBack = False
        'Set Transfer.Source = Nothing
        Set Transfer2.Source.Table = Nothing
        'Set Transfer.Destination = Nothing
        Set Transfer2.Destination.Table = Nothing
        GoTo Rewind1
    End If
    PrintTime "TryGetKeyColumns"
    
Rewind3:
    ' TODO Migrating Transfer to Transfer2
    'If Not Transfer.UnRef Is Nothing And Transfer.ValuePairs.Count = 0 Then
    '    Transfer.TryLoadValuePairs
    'End If
    If TryMapValueColumns = False Then Exit Sub
    If GoBack Then
        GoBack = False
        ' ???
        GoTo Rewind2
    End If
    PrintTime "TryMapValueColumns"
    
    'NoRewind
    DoTransfer
    PrintTime "DoTransfer"
    
    'NoRewind
    If TryTransferPreview = False Then Exit Sub
    If GoBack Then
        GoBack = False
        GoTo Rewind3
    End If
    PrintTime "TryTransferPreview"
    
    DoPostProcessing
    PrintTime "DoPostProcessing"
    
    'NoRewind
    'TrySaveHistory
    'PrintTime "TrySaveHistory"
End Sub

Private Sub InitializeViewModels()
    'Set Transfer = New TransferInstruction
    'Transfer.SetDefaultFlags
    
    Set Transfer2 = New TransferInstruction2
    
    Set SelectTableVM = New SelectTableViewModel
End Sub

'@Obsolete
Private Function TryLoadHistory() As Boolean
End Function

Private Function CheckTablesAvailable() As Boolean
    If SelectTableVM.CanSelect = False Then
        MsgBox "Not enough tables available to start a transfer!", vbExclamation + vbOKOnly
        Exit Function
    End If
    
    CheckTablesAvailable = True
End Function

Private Function CheckIfSelectionContainsTable() As Boolean
    If Selection.ListObject Is Nothing Then
        CheckIfSelectionContainsTable = False
    End If
    
    Set OneTableSelected = Selection.ListObject
End Function

Private Function TryGetSourceOrDestination() As Boolean
    TryGetSourceOrDestination = True
    
    ' TODO CHK If we need this guard
    'If Not Transfer.Source Is Nothing And Not Transfer.Destination Is Nothing Then
    '    ' Tables already set in TransferInstruction
    '    ' TODO Check what happens if we set them, but press Back button on next dialog
    '    Exit Function
    'End If
    
    If OneTableSelected Is Nothing Then Exit Function
    
    Dim vm As SourceOrDestinationViewModel
    Set vm = New SourceOrDestinationViewModel
    Set vm.ListObject = OneTableSelected
    
    Dim View As IView
    Set View = New SourceOrDestinationView
    If View.ShowDialog(vm) Then
        If vm.IsSource Then
            'Set Transfer.Source = OneTableSelected
            Set Transfer2.Source.Table = OneTableSelected
        ElseIf vm.IsDestination Then
            'Set Transfer.Destination = OneTableSelected
            Set Transfer2.Destination.Table = OneTableSelected
        Else
            TryGetSourceOrDestination = False
        End If
    Else
        TryGetSourceOrDestination = False
    End If
End Function

Private Function TryGetSecondTable() As Boolean
    Set SelectTableVM.ActiveTable = OneTableSelected
        
    If Not Transfer2.Source.Table Is Nothing And Not Transfer2.Destination.Table Is Nothing Then
        ' Tables already set in TransferInstruction
        ' TODO Check what happens if we set them, but press Back button on next dialog
        ' Might need to implement arg Optional Force as Boolean
        TryGetSecondTable = True
        Exit Function
    End If
    
    If TrySelectTable(Nothing, SelectTableVM) Then
        If Transfer2.Source.Table Is Nothing Then
            'Set Transfer.Source = SelectTableVM.SelectedTable
            Set Transfer2.Source.Table = SelectTableVM.SelectedTable
        ElseIf Transfer2.Destination.Table Is Nothing Then
            'Set Transfer.Destination = SelectTableVM.SelectedTable
            Set Transfer2.Destination.Table = SelectTableVM.SelectedTable
        Else
            Debug.Assert False
        End If
        TryGetSecondTable = True
    End If
End Function

Private Function TryGetKeyColumns() As Boolean
    TryGetKeyColumns = True
    
    Dim vm As KeyMapperViewModel
    Set vm = New KeyMapperViewModel
    Set vm.LHSTable = Transfer2.Source.Table
    Set vm.RHSTable = Transfer2.Destination.Table
    
    Dim frm As IView
    Set frm = New KeyMapperView
    
    If frm.ShowDialog(vm) Then
        If vm.GoBack Then
            GoBack = True
        ElseIf vm.IsValid Then
            If Transfer2.Source.Table Is Nothing Or Transfer2.Destination.Table Is Nothing Then
                ' TODO Migrating Transfer to Transfer2
                'CollectionHelpers.CollectionClear Transfer.ValuePairs
            ElseIf Transfer2.Source.Table <> vm.LHSTable Or Transfer2.Destination.Table <> vm.RHSTable Then
                ' TODO Migrating Transfer to Transfer2
                'CollectionHelpers.CollectionClear Transfer.ValuePairs
            End If
            
            'Set Transfer.Source = vm.LHSTable
            'Set Transfer.Destination = vm.RHSTable
            'Set Transfer.SourceKey = vm.LHSKeyColumn
            'Set Transfer.DestinationKey = vm.RHSKeyColumn
            
            Set Transfer2.Source.Table = vm.LHSTable
            Set Transfer2.Destination.Table = vm.RHSTable
            Transfer2.Source.KeyColumnName = vm.LHSKeyColumn.Name
            Transfer2.Destination.KeyColumnName = vm.RHSKeyColumn.Name
            
            'If vm.AppendNewKeys Then
            '    Transfer.Flags = AddFlag(Transfer.Flags, AppendUnmapped)
            'Else
            '    Transfer.Flags = RemoveFlag(Transfer.Flags, AppendUnmapped)
            'End If
            
            'If vm.RemoveOrphanKeys Then
            '    Transfer.Flags = AddFlag(Transfer.Flags, RemoveUnmapped)
            'Else
            '    Transfer.Flags = RemoveFlag(Transfer.Flags, RemoveUnmapped)
            'End If
            
            TryGetKeyColumns = True
        Else
            MsgBox "Invalid KeyMapperViewModel"
        End If
    Else
        TryGetKeyColumns = False
    End If
End Function

Private Function TryMapValueColumns() As Boolean
    TryMapValueColumns = True
    
    Dim vm As ValueMapperViewModel
    Set vm = New ValueMapperViewModel
    
    Set vm.LHS = Transfer2.Source.Table
    Set vm.RHS = Transfer2.Destination.Table
    Set vm.KeyColumnLHS = Transfer2.Source.KeyColumn
    Set vm.KeyColumnRHS = Transfer2.Destination.KeyColumn
    'vm.Flags = Transfer.Flags
    
    'vm.LoadFromTransferInstruction Transfer
    
    Dim frm As IView
    Set frm = New ValueMapperView
    
    If frm.ShowDialog(vm) Then
        If vm.GoBack Then
            GoBack = True
        Else
            'Set Transfer.ValuePairs = vm.checked ' Collection<ColumnPair>
            'Transfer.Flags = vm.Flags
            
            UpdateTransferInstruction Transfer2, vm.checked
            
            TryMapValueColumns = True
        End If
    Else
        TryMapValueColumns = False
    End If
End Function

Private Sub DoTransfer()
    Dim timeStart As Double
    timeStart = Timer()
    
    Transfer2.Evaluate
    
    Dim timeTaken As Double
    timeTaken = Timer() - timeStart
    
    'Dim timeStr As String
    'timeStr = Format$(timeTaken, "0.00")
    
    'Dim completionMessage As String
    'completionMessage = "Table transfer complete." & vbCrLf & "Time taken: " & timeStr & " second(s)"
    
    'MsgBox completionMessage, vbInformation + vbOKOnly, "Table Transfer Tool"
End Sub

Private Function TryTransferPreview() As Boolean
    Dim View As IView2
    Set View = TransferDeltasView
    
    Dim ViewModel As TransferDeltasViewModel
    Set ViewModel = New TransferDeltasViewModel
    ViewModel.Load Transfer2.TransferDeltas
    
    Select Case View.ShowDialog(ViewModel)
        Case ViewResultEnum.vrCancel
            TryTransferPreview = False
        Case ViewResultEnum.vrBack
            TryTransferPreview = True
            GoBack = True
        Case ViewResultEnum.vrNext
            TryTransferPreview = True
            Transfer2.Commit CommitterFactory.FullColumn
        Case ViewResultEnum.vrFinish
            TryTransferPreview = True
            Transfer2.Commit CommitterFactory.FullColumn
    End Select
End Function

Private Sub DoPostProcessing()
    Transfer2.PostProcess New RemoveHighlighting
    Transfer2.PostProcess HighlightChanges.Create(4, RGB(226, 239, 218)) ' unchanged
    Transfer2.PostProcess HighlightChanges.Create(2, RGB(204, 255, 153)) ' 0->A new
End Sub

Private Sub TrySaveHistory()
    'TransferHistorySerializer.TrySave Transfer
End Sub
