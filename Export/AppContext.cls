VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.AppContext"
Option Explicit
Implements IAppContext

Private Type TState
    Transfer As TransferInstruction
    Settings As TransferSettings
    BindingManager As BindingManager
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Settings = New TransferSettings
    Set This.Transfer = New TransferInstruction
    Set This.BindingManager = New BindingManager
End Sub

Private Property Get IAppContext_BindingManager() As BindingManager
    Set IAppContext_BindingManager = This.BindingManager
End Property

Public Sub Start()
    'On Error GoTo ErrorHandler

'DEBUG:
    'Set This.Transfer = GetDebugTransfer
    'GoTo RewindValueMapper

RewindReloadRecent:
    Select Case TryReloadRecent
        Case vrNext
            GoTo RewindTransferDeltas
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
    End Select
    
RewindTransferDeltas:
    Select Case TryTransferDeltas
        Case vrStart
            GoTo RewindKeyMapper
        Case vrBack
            GoTo RewindValueMapper
        Case vrCancel, vrNext
            GoTo FinishEarly
    End Select
    
FinishEarly:
    SaveTransferHistory
    Exit Sub
    
ErrorHandler:
    MsgBox ERR_MSG_GENERIC, vbCritical + vbOKOnly, MSG_CAPTION
End Sub

Private Function TryReloadRecent() As TtViewResult
    ' DEBUG
    Exit Function
    
    This.Settings.LoadMRU
    
    Dim RecentTransfers As Collection
    Set RecentTransfers = This.Settings.RecentTransfers.ToCollection
    
    Dim RecentTransfer As Variant
    For Each RecentTransfer In RecentTransfers
        Dim SerialString As String
        SerialString = Base64toString(RecentTransfer)
        
        Dim TestTransfer As TransferInstruction
        If TransferInstructionSerializer.TryDeserialize(SerialString, TestTransfer) Then
            Set This.Transfer = TestTransfer
            TryReloadRecent = vrNext
            Exit Function
        End If
    Next RecentTransfer
End Function

Private Function TryTablePicker() As TtViewResult
    Dim ViewModel As TablePickerViewModel
    Set ViewModel = New TablePickerViewModel
    ViewModel.Load
    
    Set ViewModel.SourceTable = This.Transfer.Source.Table
    Set ViewModel.DestinationTable = This.Transfer.Destination.Table
    
    Dim View As IView
    Set View = TablePickerView.Create(Me, ViewModel)
    
    Dim Result As TtViewResult
    Result = View.ShowDialog
  
    Select Case Result
        Case vrNext
            Set This.Transfer.Source.Table = ViewModel.SourceTable
            Set This.Transfer.Destination.Table = ViewModel.DestinationTable
    End Select
    
    TryTablePicker = Result
End Function

Private Function TryKeyMapper() As TtViewResult
    Dim ViewModel As KeyMapperViewModel
    Set ViewModel = New KeyMapperViewModel
    ViewModel.Load This.Transfer.Source.Table, This.Transfer.Destination.Table
    
    Dim ListObjectSettings As ISettings
    Set ListObjectSettings = This.Settings.GetListObjectSettings(This.Transfer.Source.Table)
    ViewModel.Source.TrySelect ListObjectSettings.GetSetting(LAST_KEY_USED)
    
    Set ListObjectSettings = This.Settings.GetListObjectSettings(This.Transfer.Destination.Table)
    ViewModel.Destination.TrySelect ListObjectSettings.GetSetting(LAST_KEY_USED)
    
    ViewModel.TryEvaluateMatch
    
    Dim View As IView
    Set View = KeyMapperView.Create(Me, ViewModel)
    
    Dim Result As TtViewResult
    Result = View.ShowDialog()
    
    Select Case Result
        Case vrNext
            This.Transfer.Source.KeyColumnName = ViewModel.Source.Selected.Name
            This.Transfer.Destination.KeyColumnName = ViewModel.Destination.Selected.Name
    End Select
    
    TryKeyMapper = Result
End Function

Private Function TryValueMapper() As TtViewResult
    Dim ViewModel As ValueMapperViewModel
    Set ViewModel = New ValueMapperViewModel
    With ViewModel.Source
        .AllowEmpty = True
        .AllowError = False
        .AllowFormula = True
        .Load This.Transfer.Source.Table
        .SetKeyColumn This.Transfer.Source.KeyColumnName
    End With
    
    With ViewModel.Destination
        .AllowEmpty = True
        .AllowError = True
        .AllowFormula = False
        .Load This.Transfer.Destination.Table
        .SetKeyColumn This.Transfer.Destination.KeyColumnName
    End With
    
    Dim View As IView
    Set View = ValueMapperView.Create(Me, ViewModel)
    
    Dim Result As TtViewResult
    Result = View.ShowDialog()
    
    Select Case Result
        Case vrNext
            With This.Transfer
                .Source.Load .Source.Table, .Source.KeyColumnName, ViewModel.Mapped.GetArray(ttSource)
                .Destination.Load .Destination.Table, .Destination.KeyColumnName, ViewModel.Mapped.GetArray(ttDestination)
            End With
    End Select
    
    TryValueMapper = Result
End Function

Private Function TryTransferDeltas() As TtViewResult
    This.Transfer.Evaluate
    
    Dim ViewModel As DeltasPreviewViewModel
    Set ViewModel = New DeltasPreviewViewModel
    ViewModel.Load This.Transfer.TransferDeltas, This.Transfer.Source.Table, This.Transfer.Destination.Table
    
    Dim View As IView
    Set View = DeltasPreviewView.Create(Me, ViewModel)
    
    Dim Result As TtViewResult
    Result = View.ShowDialog()
    
    Select Case Result
        Case vrNext
            This.Transfer.Commit CommitterFactory.FullColumn
            DoPostProcessing
    End Select
    
    TryTransferDeltas = Result
End Function

Private Sub DoPostProcessing()
    If IsEmpty(This.Transfer.Deltas) Then Exit Sub
    
    With This.Transfer
        .PostProcess PostProcessorFactory.RemoveHighlighting
        .PostProcess HighlightChanges.Create(ttValueChanged, COLOR_GREEN_DARK)
        .PostProcess HighlightChanges.Create(ttValueReplacesBlank, COLOR_GREEN_LIGHT)
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
