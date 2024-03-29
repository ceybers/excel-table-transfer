Attribute VB_Name = "TransferHistorySerializer"
'@IgnoreModule
'@Folder "Model.TransferHistory"
Option Explicit

Private Const WORKSHEET_NAME As String = "CAETransferTableHistory"
Private Const RANGE_TO_REMOVE As String = "A:D"

Public Function TryLoad(ByRef TransferInstructionUnref As TransferInstructionUnref) As Boolean
    Dim Worksheet As Worksheet
    
    If TryGetWorksheet(ThisWorkbook, WORKSHEET_NAME, Worksheet) = False Then
        Exit Function
    End If
    
    Dim Range As Range
    Set Range = Worksheet.Range("A1")
    
    Set TransferInstructionUnref = New TransferInstructionUnref
    If TransferInstructionUnref.LoadFromRange(Range) = False Then
        Exit Function
    End If
    
    TryLoad = True
End Function

Public Function TrySave(ByVal TransferInstruction As TransferInstruction) As Boolean
    Debug.Assert Not TransferInstruction Is Nothing
    
    Dim Worksheet As Worksheet
    
    If TryGetWorksheet(ThisWorkbook, WORKSHEET_NAME, Worksheet) = False Then
        Set Worksheet = CreateNewHistoryWorksheet
    End If
    
    Worksheet.Range(RANGE_TO_REMOVE).Clear
    
    Dim Range As Range
    Set Range = Worksheet.Range("A1")
    TransferInstruction.SaveToRange Range
    
    TrySave = True
End Function

Private Function CreateNewHistoryWorksheet() As Worksheet
    Dim CurrentWorksheet As Worksheet
    Set CurrentWorksheet = ActiveSheet
    
    Dim Result As Worksheet
    Set Result = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets.Item(ActiveWorkbook.Worksheets.Count))
    With Result
        .Name = WORKSHEET_NAME
        .Visible = xlSheetVeryHidden
    End With
    
    CurrentWorksheet.Activate
    
    Set CreateNewHistoryWorksheet = Result
End Function
