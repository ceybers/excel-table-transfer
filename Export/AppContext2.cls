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

Private Const LAST_KEY_USED As String = "LastKeyUsed"

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

RewindTablePicker:
    Select Case TryTablePicker
        Case vrCancel
            Exit Sub
    End Select
    
RewindKeyMapper:
    Select Case TryKeyMapper
        Case vrBack
            GoTo RewindTablePicker
        Case vrCancel
            Exit Sub
    End Select
    
RewindValueMapper:
    Select Case TryValueMapper
        Case vrBack
            GoTo RewindKeyMapper
        Case vrCancel
            Exit Sub
        'Case vrFinish
        '    GoTo FinishEarly
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

Private Function TryTablePicker() As ViewResult
    Dim ViewModel As TablePickerViewModel
    Set ViewModel = New TablePickerViewModel
    ViewModel.Load
    
    Set ViewModel.SourceTable = This.Transfer.Source.Table
    Set ViewModel.DestinationTable = This.Transfer.Destination.Table
    
    Dim View As IView2
    Set View = New TablePickerView
    
    Dim Result As ViewResult
    Result = View.ShowDialog(ViewModel)
  
    Select Case Result
        Case vrNext
            Set This.Transfer.Source.Table = ViewModel.SourceTable
            Set This.Transfer.Destination.Table = ViewModel.DestinationTable
    End Select
    
    TryTablePicker = Result
End Function

Private Function TryKeyMapper() As ViewResult
    Dim ViewModel As KeyMapperViewModel
    Set ViewModel = New KeyMapperViewModel
    ViewModel.LoadFromTransferInstruction2 This.Transfer
    
    Dim ListObjectSettings As ISettings
    Set ListObjectSettings = This.Settings.GetListObjectSettings(This.Transfer.Source.Table)
    This.Transfer.Source.KeyColumnName = ListObjectSettings.GetSetting(LAST_KEY_USED)
    
    Set ListObjectSettings = This.Settings.GetListObjectSettings(This.Transfer.Destination.Table)
    This.Transfer.Destination.KeyColumnName = ListObjectSettings.GetSetting(LAST_KEY_USED)
    
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
        'Case vrFinish
            'UpdateTransferInstruction This.Transfer, ViewModel.CheckedValuePairs
            'This.Transfer.Evaluate
            'If IsEmpty(This.Transfer.Deltas) Then
            '    Result = vrNext
            'Else
            '    This.Transfer.Commit CommitterFactory.FullColumn
            '    DoPostProcessing
            'End If
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
    If IsEmpty(This.Transfer.Deltas) Then Exit Sub
    
    With This.Transfer
        .PostProcess PostProcessorFactory.RemoveHighlighting
        .PostProcess HighlightChanges.Create(4, RGB(226, 239, 218)) ' unchanged
        .PostProcess HighlightChanges.Create(2, RGB(204, 255, 153)) ' 0->A new
        .PostProcess PostProcessorFactory.TransferColumnWidths
        .PostProcess PostProcessorFactory.TransferNumberFormat
    End With
End Sub

Private Sub SaveTransferHistory()
    If IsEmpty(This.Transfer.Deltas) Then Exit Sub
    
    Dim SerialString As String
    SerialString = TransferInstructionSerializer.Serialize(This.Transfer)
    
    Dim Base64Serial As String
    Base64Serial = StringtoBase64(SerialString)
    
    This.Settings.RecentTransfers.Add Base64Serial
    This.Settings.SaveMRU
    
    Dim ListObjectSettings As ISettings
    Set ListObjectSettings = This.Settings.GetListObjectSettings(This.Transfer.Source.Table)
    ListObjectSettings.SetSetting LAST_KEY_USED, This.Transfer.Source.KeyColumnName
    
    Set ListObjectSettings = This.Settings.GetListObjectSettings(This.Transfer.Destination.Table)
    ListObjectSettings.SetSetting LAST_KEY_USED, This.Transfer.Destination.KeyColumnName
End Sub
