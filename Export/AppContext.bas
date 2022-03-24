Attribute VB_Name = "AppContext"
Option Explicit

Private transfer As TransferInstruction
Private vm As SelectTableViewModel

Private OneTableSelected As ListObject

Private GoBack As Boolean

Public Sub AAA_DoWork()
    'MsgBox "Welcome to table transfer wizard"
    
    InitializeViewModels
    
    If CheckTablesAvailable = False Then Exit Sub
    
Rewind1:
    CheckIfSelectionContainsTable
    
    If TryGetSourceOrDestination = False Then Exit Sub
    
    If TryGetSecondTable = False Then Exit Sub
    
Rewind2:
    If TryGetKeyColumns = False Then Exit Sub
    If GoBack Then
        GoBack = False
        Set transfer.Source = Nothing
        Set transfer.Destination = Nothing
        GoTo Rewind1
    End If

Rewind3:
    If TryMapValueColumns = False Then Exit Sub
    If GoBack Then
        GoBack = False
        ' ???
        GoTo Rewind2
    End If
    
'NoRewind
    DoTransfer
    
'NoRewind
    TrySaveHistory
End Sub

Private Sub InitializeViewModels()
    Set transfer = New TransferInstruction
    transfer.SetDefaultFlags
    
    Set vm = New SelectTableViewModel
End Sub

Private Function CheckTablesAvailable() As Boolean
    If vm.Tables.Count < 2 Then
        MsgBox "Not enough tables available to start a transfer!", vbExclamation + vbOKOnly
        Exit Function
    End If
    CheckTablesAvailable = True
End Function

Private Function CheckIfSelectionContainsTable() As Boolean
    If Selection.ListObject Is Nothing Then
    
    End If
    
    Set OneTableSelected = Selection.ListObject
End Function

Private Function TryGetSourceOrDestination() As Boolean
    TryGetSourceOrDestination = True
    
    If OneTableSelected Is Nothing Then Exit Function
    
    Dim vm As SourceOrDestinationViewModel
    Set vm = New SourceOrDestinationViewModel
    Set vm.ListObject = OneTableSelected
    
    Dim view As IView
    Set view = New SourceOrDestinationView
    If view.ShowDialog(vm) Then
        If vm.IsSource Then
            Set transfer.Source = OneTableSelected
        ElseIf vm.IsDestination Then
            Set transfer.Destination = OneTableSelected
        Else
            TryGetSourceOrDestination = False
        End If
    Else
        TryGetSourceOrDestination = False
    End If
End Function

Private Function TryGetSecondTable() As Boolean
    Dim vm As SelectTableViewModel
    Set vm = New SelectTableViewModel
    Set vm.ActiveTable = OneTableSelected
    
    Dim frm As IView
    Set frm = New SelectTableView
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            If transfer.Source Is Nothing Then
                Set transfer.Source = vm.SelectedTable
            Else
                Set transfer.Destination = vm.SelectedTable
            End If
            TryGetSecondTable = True
        End If
    End If
End Function

Private Function TryGetKeyColumns() As Boolean
    TryGetKeyColumns = True
    
    Dim vm As KeyMapperViewModel
    Set vm = New KeyMapperViewModel
    Set vm.LHSTable = transfer.Source
    Set vm.RHSTable = transfer.Destination
    
    Dim frm As IView
    Set frm = New KeyMapperView
    
    If frm.ShowDialog(vm) Then
        If vm.GoBack Then
            GoBack = True
        ElseIf vm.IsValid Then
            Set transfer.Source = vm.LHSTable
            Set transfer.Destination = vm.RHSTable
            Set transfer.SourceKey = vm.LHSKeyColumn
            Set transfer.DestinationKey = vm.RHSKeyColumn
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
    
    Set vm.lhs = transfer.Source
    Set vm.RHS = transfer.Destination
    Set vm.KeyColumnLHS = transfer.SourceKey
    Set vm.KeyColumnRHS = transfer.DestinationKey
    vm.Flags = transfer.Flags
    
    vm.LoadFromTransferInstruction transfer
    
    Dim frm As IView
    Set frm = New ValueMapperView
    
    If frm.ShowDialog(vm) Then
        If vm.GoBack Then
            GoBack = True
        Else
            Set transfer.ValuePairs = vm.checked
            transfer.Flags = vm.Flags
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
    transfer.transfer
    timeTaken = Timer() - timeStart
    
    Dim timeStr As String
    timeStr = Format(timeTaken, "0.00")
    
    Dim completionMessage As String
    completionMessage = "Table transfer complete." & vbCrLf & "Time taken: " & timeStr & " second(s)"
    
    MsgBox completionMessage, vbInformation + vbOKOnly, "Table Transfer Tool"
End Sub

Private Sub TrySaveHistory()
    If HasFlag(transfer.Flags, saveToHistory) Then
        Dim history As TransferHistoryViewModel
        Set history = New TransferHistoryViewModel
        If history.HasHistory = False Then
            history.Create
        End If
        history.Refresh
        history.Add transfer
        history.Save
    End If
End Sub
