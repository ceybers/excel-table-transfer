Attribute VB_Name = "TransferHistorySerializer"
'@IgnoreModule
'@Folder "ZZZTransferHistory"
Option Explicit

Private Const WORKSHEET_NAME As String = "CAETransferTableHistory"
Private Const RANGE_TO_REMOVE As String = "A:D"

Public Function TryLoad(ByRef tiUr As TransferInstructionUnref) As Boolean
    Dim ws As Worksheet
    
    If TryGetWorksheet(ThisWorkbook, WORKSHEET_NAME, ws) = False Then
        Exit Function
    End If
    
    Dim rng
    Set rng = ws.Range("A1")
    
    Set tiUr = New TransferInstructionUnref
    If tiUr.LoadFromRange(rng) = False Then
        Exit Function
    End If
    
    TryLoad = True
End Function

Public Function TrySave(ByVal ti As TransferInstruction) As Boolean
    Debug.Assert Not ti Is Nothing
    
    Dim ws As Worksheet
    
    If TryGetWorksheet(ThisWorkbook, WORKSHEET_NAME, ws) = False Then
        Dim curWS As Worksheet
        Set curWS = ActiveSheet
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = WORKSHEET_NAME
        ws.Visible = xlSheetVeryHidden
        curWS.Activate
    End If
    
    ws.Range(RANGE_TO_REMOVE).Clear
    
    Dim rng
    Set rng = ws.Range("A1")
    ti.SaveToRange rng
    
    TrySave = True
End Function

