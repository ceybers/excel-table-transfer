VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferHistoryViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule
'@Folder "Model.TransferHistory"
Option Explicit
Implements IViewModel

Private Const WORKSHEET_NAME As String = "CAETransferTableHistory"

Public Items As Collection
Public Worksheet As Worksheet
Public SelectedInstruction As TransferInstruction

Private Sub Class_Initialize()
    Set Items = New Collection
    If HasHistory Then
        'Set Worksheet = WorksheetHelpers.TryGetWorksheet(WORKSHEET_NAME)
        TryGetWorksheet ThisWorkbook, WORKSHEET_NAME, Worksheet
    End If
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = Not SelectedInstruction Is Nothing
End Function

Public Function HasHistory() As Boolean
    HasHistory = WorksheetHelpers.WorksheetExists(ActiveWorkbook, WORKSHEET_NAME)
End Function

Public Sub Add(ByVal TransferInstruction As TransferInstruction)
    Debug.Assert Not Worksheet Is Nothing
    Items.Add TransferInstruction
End Sub

'@Obsolete
Public Sub Save()
    'Worksheet.UsedRange.EntireRow.Delete
    'SaveTransferInstructionsFromWorksheet Items, Worksheet
End Sub

'@Obsolete
Public Sub Refresh()
    'Dim Result As Variant
    'Debug.Assert Not Worksheet Is Nothing
    'CollectionHelpers.CollectionClear Items
    'If Not Worksheet Is Nothing Then
    '    Set Result = LoadTransferInstructionsFromWorksheet(Worksheet)
    '    If Not Result Is Nothing Then
    '        Set Items = Result
    '    End If
    'End If
End Sub

Public Sub Clear()
    Debug.Assert Not Worksheet Is Nothing
    Worksheet.UsedRange.EntireRow.Delete
    CollectionHelpers.CollectionClear Items
End Sub

Public Sub Create()
    If HasHistory Then
        Exit Sub
    End If

    Dim curWS As Worksheet
    Set curWS = ActiveSheet
    Set Worksheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets.Item(ActiveWorkbook.Worksheets.Count))
    Worksheet.Name = WORKSHEET_NAME
    Worksheet.Visible = xlSheetVeryHidden
    curWS.Activate
End Sub

Public Sub Remove()
    If Not HasHistory Then
        Exit Sub
    End If
    
    ' TODO Const
    If vbYes = MsgBox("Are you sure you want to remove Table Transfer History?", vbExclamation + vbYesNo + vbDefaultButton2, "Remove Table Transfer History") Then
        Application.DisplayAlerts = False
        Worksheet.Visible = xlSheetHidden
        Worksheet.Delete
        Application.DisplayAlerts = True
        Set Worksheet = Nothing
    End If
End Sub


