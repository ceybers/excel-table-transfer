Attribute VB_Name = "AppContext"
'@IgnoreModule EmptyIfBlock
'@Folder "TableTransfer"
Option Explicit

Private transfer As TransferInstruction
Private SelectTableVM As SelectTableViewModel

Private OneTableSelected As ListObject

Private GoBack As Boolean

Public Sub TransferTable()
    'MsgBox "Welcome to table transfer wizard"
    
    InitializeViewModels
    
    ' DEBUG
    Set transfer.Source = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set transfer.Destination = ThisWorkbook.Worksheets(1).ListObjects(2)
    
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
    
    Set SelectTableVM = New SelectTableViewModel
End Sub

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
    
    If Not transfer.Source Is Nothing And Not transfer.Destination Is Nothing Then
        ' Tables already set in TransferInstruction
        ' TODO Check what happens if we set them, but press Back button on next dialog
        Exit Function
    End If
    
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
    Set SelectTableVM.ActiveTable = OneTableSelected
        
    If Not transfer.Source Is Nothing And Not transfer.Destination Is Nothing Then
        ' Tables already set in TransferInstruction
        ' TODO Check what happens if we set them, but press Back button on next dialog
        ' Might need to implement arg Optional Force as Boolean
        TryGetSecondTable = True
        Exit Function
    End If
    
    If TrySelectTable(Nothing, SelectTableVM) Then
        If transfer.Source Is Nothing Then
            Set transfer.Source = SelectTableVM.SelectedTable
        ElseIf transfer.Destination Is Nothing Then
            Set transfer.Destination = SelectTableVM.SelectedTable
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
            
            If vm.AppendNewKeys Then
                transfer.Flags = AddFlag(transfer.Flags, AppendUnmapped)
            Else
                transfer.Flags = RemoveFlag(transfer.Flags, AppendUnmapped)
            End If
            
            If vm.RemoveOrphanKeys Then
                transfer.Flags = AddFlag(transfer.Flags, RemoveUnmapped)
            Else
                transfer.Flags = RemoveFlag(transfer.Flags, RemoveUnmapped)
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
    
    Set vm.LHS = transfer.Source
    Set vm.rhs = transfer.Destination
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
    timeStr = Format$(timeTaken, "0.00")
    
    Dim completionMessage As String
    completionMessage = "Table transfer complete." & vbCrLf & "Time taken: " & timeStr & " second(s)"
    
    MsgBox completionMessage, vbInformation + vbOKOnly, "Table Transfer Tool"
End Sub

Private Sub TrySaveHistory()
    If HasFlag(transfer.Flags, SaveToHistory) Then
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
