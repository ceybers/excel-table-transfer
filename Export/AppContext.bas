Attribute VB_Name = "AppContext"
'@IgnoreModule EmptyIfBlock
'@Folder "MVVM.AppContext"
Option Explicit

Private Transfer As TransferInstruction
Private SelectTableVM As SelectTableViewModel

Private OneTableSelected As ListObject

Private GoBack As Boolean

Public Sub PrintTime(ByVal message As String, Optional ByVal Reset As Boolean)
    Static startTime As Double
    If Reset Or (startTime = 0) Then
        startTime = Timer()
    End If
    Debug.Print message & " " & (Timer() - startTime)
End Sub

Public Sub TransferTable()
Attribute TransferTable.VB_ProcData.VB_Invoke_Func = "e\n14"
    PrintTime "Start", True
       
    'MsgBox "Welcome to table transfer wizard"
    
    InitializeViewModels
    
    If TryLoadHistory Then
        GoTo Rewind3
    End If
    
    ' DEBUG
    'Set Transfer.Source = ThisWorkbook.Worksheets(1).ListObjects(1)
    'Set Transfer.Destination = ThisWorkbook.Worksheets(1).ListObjects(2)
    
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
        Set Transfer.Source = Nothing
        Set Transfer.Destination = Nothing
        GoTo Rewind1
    End If
    PrintTime "TryGetKeyColumns"
    
Rewind3:
    If Not Transfer.UnRef Is Nothing And Transfer.ValuePairs.Count = 0 Then
        Transfer.TryLoadValuePairs
    End If
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
    TrySaveHistory
    PrintTime "TrySaveHistory"
End Sub

Private Sub InitializeViewModels()
    Set Transfer = New TransferInstruction
    Transfer.SetDefaultFlags
    
    Set SelectTableVM = New SelectTableViewModel
End Sub

Private Function TryLoadHistory() As Boolean
    TryLoadHistory = True
    Dim tiUr As TransferInstructionUnref
    If TransferHistorySerializer.TryLoad(tiUr) Then
        Set Transfer.UnRef = tiUr
        Transfer.LoadFlags
        If Transfer.TryLoadTables = False Then
            TryLoadHistory = False
        End If
        If Transfer.TryLoadKeyColumns = False Then
            TryLoadHistory = False
        End If
        'Debug.Print Transfer.TryLoadValuePairs
    Else
        TryLoadHistory = False
    End If
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
    
    If Not Transfer.Source Is Nothing And Not Transfer.Destination Is Nothing Then
        ' Tables already set in TransferInstruction
        ' TODO Check what happens if we set them, but press Back button on next dialog
        Exit Function
    End If
    
    If OneTableSelected Is Nothing Then Exit Function
    
    Dim vm As SourceOrDestinationViewModel
    Set vm = New SourceOrDestinationViewModel
    Set vm.ListObject = OneTableSelected
    
    Dim View As IView
    Set View = New SourceOrDestinationView
    If View.ShowDialog(vm) Then
        If vm.IsSource Then
            Set Transfer.Source = OneTableSelected
        ElseIf vm.IsDestination Then
            Set Transfer.Destination = OneTableSelected
        Else
            TryGetSourceOrDestination = False
        End If
    Else
        TryGetSourceOrDestination = False
    End If
End Function

Private Function TryGetSecondTable() As Boolean
    Set SelectTableVM.ActiveTable = OneTableSelected
        
    If Not Transfer.Source Is Nothing And Not Transfer.Destination Is Nothing Then
        ' Tables already set in TransferInstruction
        ' TODO Check what happens if we set them, but press Back button on next dialog
        ' Might need to implement arg Optional Force as Boolean
        TryGetSecondTable = True
        Exit Function
    End If
    
    If TrySelectTable(Nothing, SelectTableVM) Then
        If Transfer.Source Is Nothing Then
            Set Transfer.Source = SelectTableVM.SelectedTable
        ElseIf Transfer.Destination Is Nothing Then
            Set Transfer.Destination = SelectTableVM.SelectedTable
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
    Set vm.LHSTable = Transfer.Source
    Set vm.RHSTable = Transfer.Destination
    
    Dim frm As IView
    Set frm = New KeyMapperView
    
    If frm.ShowDialog(vm) Then
        If vm.GoBack Then
            GoBack = True
        ElseIf vm.IsValid Then
            If Transfer.Source Is Nothing Or Transfer.Destination Is Nothing Then
                CollectionHelpers.CollectionClear Transfer.ValuePairs
            ElseIf Transfer.Source <> vm.LHSTable Or Transfer.Destination <> vm.RHSTable Then
                CollectionHelpers.CollectionClear Transfer.ValuePairs
            End If
            
            Set Transfer.Source = vm.LHSTable
            Set Transfer.Destination = vm.RHSTable
            Set Transfer.SourceKey = vm.LHSKeyColumn
            Set Transfer.DestinationKey = vm.RHSKeyColumn
            
            If vm.AppendNewKeys Then
                Transfer.Flags = AddFlag(Transfer.Flags, AppendUnmapped)
            Else
                Transfer.Flags = RemoveFlag(Transfer.Flags, AppendUnmapped)
            End If
            
            If vm.RemoveOrphanKeys Then
                Transfer.Flags = AddFlag(Transfer.Flags, RemoveUnmapped)
            Else
                Transfer.Flags = RemoveFlag(Transfer.Flags, RemoveUnmapped)
            End If
            
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
    
    Set vm.LHS = Transfer.Source
    Set vm.RHS = Transfer.Destination
    Set vm.KeyColumnLHS = Transfer.SourceKey
    Set vm.KeyColumnRHS = Transfer.DestinationKey
    vm.Flags = Transfer.Flags
    
    vm.LoadFromTransferInstruction Transfer
    
    Dim frm As IView
    Set frm = New ValueMapperView
    
    If frm.ShowDialog(vm) Then
        If vm.GoBack Then
            GoBack = True
        Else
            Set Transfer.ValuePairs = vm.checked
            Transfer.Flags = vm.Flags
            TryMapValueColumns = True
        End If
    Else
        TryMapValueColumns = False
    End If
End Function

Private Sub DoTransfer()
    Dim timeStart As Long
    Dim timeTaken As Long
    timeStart = Timer()
    Transfer.Transfer
    timeTaken = Timer() - timeStart
    
    Dim timeStr As String
    timeStr = Format$(timeTaken, "0.00")
    
    Dim completionMessage As String
    completionMessage = "Table transfer complete." & vbCrLf & "Time taken: " & timeStr & " second(s)"
    
    MsgBox completionMessage, vbInformation + vbOKOnly, "Table Transfer Tool"
End Sub

Private Sub TrySaveHistory()
    TransferHistorySerializer.TrySave Transfer
    
    'If HasFlag(transfer.Flags, SaveToHistory) Then
    'Dim history As TransferHistoryViewModel
    'Set history = New TransferHistoryViewModel
    'If history.HasHistory = False Then
    '    history.Create
    'End If
    'history.Refresh
    'history.Add transfer
    'history.Save
    'End If
End Sub
