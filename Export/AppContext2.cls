VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM2.AppContext"
Option Explicit

Private Const MSG_CAPTION As String = "Table Transfer Tool"
Private Const MSG_NOT_ENOUGH_TABLES As String = "Not enough tables available to start a transfer!"

Private Type TState
    Transfer As TransferInstruction2
    Settings As TransferSettings
End Type
Private This As TState

Public Sub Start()
    Set This.Settings = New TransferSettings
    Set This.Transfer = New TransferInstruction2
    
'DEBUG:
    'Set This.Transfer = GetDebugTransfer
    'GoTo RewindValueMapper

RewindReloadRecent:
    Select Case TryReloadRecent
        Case vrNext
            GoTo RewindValueMapper
    End Select
    
RewindSourceOrDestination:
    Select Case TrySourceOrDestination
        Case vrCancel
            Exit Sub
    End Select

RewindSelectTable:
    Select Case TrySelectTable
        Case vrCancel
            Exit Sub
    End Select
    
RewindKeyMapper:
    Select Case TryKeyMapper
        'Case vrStart
        '    GoTo RewindSourceOrDestination
        Case vrBack
            GoTo RewindSelectTable
        Case vrCancel
            Exit Sub
    End Select
    
RewindValueMapper:
    Select Case TryValueMapper
        Case vrBack
            GoTo RewindKeyMapper
        Case vrCancel
            Exit Sub
        Case vrFinish
            GoTo FinishEarly
    End Select
    
RewindTransferDeltas:
    Select Case TryTransferDeltas
        Case vrStart
            GoTo RewindKeyMapper
        Case vrBack
            GoTo RewindValueMapper
        Case vrCancel, vrFinish
            GoTo FinishEarly
    End Select
    
FinishEarly:
    SaveTransferHistory
End Sub

Private Function TryReloadRecent() As ViewResult
    This.Settings.LoadMRU
    
    Dim RecentTransfers As Collection
    Set RecentTransfers = This.Settings.RecentTransfers.ToCollection
    
    Dim RecentTransfer As Variant
    For Each RecentTransfer In RecentTransfers
        Dim SerialString As String
        SerialString = Base64toString(RecentTransfer)
        
        Dim TestTransfer As TransferInstruction2
        If TransferInstructionSerializer.TryDeserialize(SerialString, TestTransfer) Then
            Set This.Transfer = TestTransfer
            TryReloadRecent = vrNext
            Exit Function
        End If
    Next RecentTransfer
End Function

Private Function TrySourceOrDestination() As ViewResult
    Dim ViewModel As SourceOrDestinationViewModel
    Set ViewModel = New SourceOrDestinationViewModel
    
    Dim View As IView2
    Set View = New SourceOrDestinationView
    
    Dim Result As ViewResult
    
    If ViewModel.CanSelect = False Then
        MsgBox MSG_NOT_ENOUGH_TABLES, vbExclamation + vbOKOnly, MSG_CAPTION
        TrySourceOrDestination = vrCancel
        Exit Function
    End If
    
    If Selection.ListObject Is Nothing Then
        TrySourceOrDestination = vrNext
        Exit Function
    End If
        
    Result = View.ShowDialog(ViewModel)
    
    Select Case Result
        Case vrNext
            If ViewModel.IsSource Then
                Set This.Transfer.Source.Table = Selection.ListObject
            ElseIf ViewModel.IsDestination Then
                Set This.Transfer.Destination.Table = Selection.ListObject
            End If
    End Select
    
    TrySourceOrDestination = Result
End Function

Private Function TrySelectTable() As ViewResult
    Dim ViewModel As SelectTableViewModel
    Set ViewModel = New SelectTableViewModel
    
    Dim View As IView2
    Set View = New SelectTableView
    
    Dim Result As ViewResult
    Result = View.ShowDialog(ViewModel)
  
    Select Case Result
        Case vrNext
            If This.Transfer.Source.Table Is Nothing Then
                Set This.Transfer.Source.Table = ViewModel.SelectedTable
            ElseIf This.Transfer.Destination.Table Is Nothing Then
                Set This.Transfer.Destination.Table = ViewModel.SelectedTable
            Else
                Err.Raise vbObjectError + 14, "TrySelectTable", "Failed to set selected table."
            End If
    End Select
    
    TrySelectTable = Result
End Function

Private Function TryKeyMapper() As ViewResult
    Dim ViewModel As KeyMapperViewModel
    Set ViewModel = New KeyMapperViewModel
    ViewModel.LoadFromTransferInstruction2 This.Transfer
    
    Dim View As IView2
    Set View = New KeyMapperView
    
    Dim Result As ViewResult
    Result = View.ShowDialog(ViewModel)
    
    Select Case Result
        Case vrNext
            ViewModel.SaveToTransferInstruction This.Transfer
    End Select
    
    TryKeyMapper = Result
End Function

Private Function TryValueMapper() As ViewResult
    Dim ViewModel As ValueMapperViewModel
    Set ViewModel = New ValueMapperViewModel
    ViewModel.LoadFromTransferInstruction2 This.Transfer
    
    Dim View As IView2
    Set View = New ValueMapperView
    
    Dim Result As ViewResult
    Result = View.ShowDialog(ViewModel)
    
    Select Case Result
        Case vrNext
            UpdateTransferInstruction This.Transfer, ViewModel.CheckedValuePairs
            This.Transfer.Evaluate
        Case vrFinish
            UpdateTransferInstruction This.Transfer, ViewModel.CheckedValuePairs
            This.Transfer.Evaluate
            This.Transfer.Commit CommitterFactory.FullColumn
            DoPostProcessing
    End Select
    
    TryValueMapper = Result
End Function

Private Function TryTransferDeltas() As ViewResult
    Dim ViewModel As TransferDeltasViewModel
    Set ViewModel = New TransferDeltasViewModel
    ViewModel.Load This.Transfer.TransferDeltas
    
    Dim View As IView2
    Set View = TransferDeltasView
    
    Dim Result As ViewResult
    Result = View.ShowDialog(ViewModel)
    
    Select Case Result
        Case vrFinish
            This.Transfer.Commit CommitterFactory.FullColumn
            DoPostProcessing
    End Select
    
    TryTransferDeltas = Result
End Function

Private Sub DoPostProcessing()
    With This.Transfer
        .PostProcess PostProcessorFactory.RemoveHighlighting
        .PostProcess HighlightChanges.Create(4, RGB(226, 239, 218)) ' unchanged
        .PostProcess HighlightChanges.Create(2, RGB(204, 255, 153)) ' 0->A new
        .PostProcess PostProcessorFactory.TransferColumnWidths
        .PostProcess PostProcessorFactory.TransferNumberFormat
    End With
End Sub

Private Sub SaveTransferHistory()
    Dim SerialString As String
    SerialString = TransferInstructionSerializer.Serialize(This.Transfer)
    
    Dim Base64Serial As String
    Base64Serial = StringtoBase64(SerialString)
    
    This.Settings.RecentTransfers.Add Base64Serial
    This.Settings.SaveMRU
End Sub
