VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferHistoryViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TransferHistory")
Option Explicit
Implements IViewModel

Private Const WORKSHEET_NAME As String = "CAETransferTableHistory"

Public Items As Collection
Public Worksheet As Worksheet
Public SelectedInstruction As TransferInstruction

Private Sub class_initialize()
    Set Items = New Collection
    If HasHistory Then
        Set Worksheet = WorksheetHelpers.GetWorksheetByName(WORKSHEET_NAME)
    End If
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = Not SelectedInstruction Is Nothing
End Function

Public Function HasHistory() As Boolean
    HasHistory = WorksheetHelpers.DoesWorksheetExist(WORKSHEET_NAME)
End Function

Public Sub Add(ByVal TransferInstruction As TransferInstruction)
    Debug.Assert Not Worksheet Is Nothing
    Items.Add TransferInstruction
End Sub

Public Sub Save()
    Worksheet.UsedRange.EntireRow.Delete
    SaveTransferInstructionsFromWorksheet Items, Worksheet
End Sub

Public Sub Refresh()
    Dim result As Variant
    'Debug.Assert Not Worksheet Is Nothing
    ClearCollection Items
    If Not Worksheet Is Nothing Then
        Set result = LoadTransferInstructionsFromWorksheet(Worksheet)
        If Not result Is Nothing Then
            Set Items = result
        End If
    End If
End Sub

Public Sub Clear()
    Debug.Assert Not Worksheet Is Nothing
    Worksheet.UsedRange.EntireRow.Delete
    ClearCollection Items
End Sub

Public Sub Create()
    If HasHistory Then
        Exit Sub
    End If

    Dim curWS As Worksheet
    Set curWS = ActiveSheet
    Set Worksheet = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
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

Public Sub InitializeListView(ByVal lv As ListView)
    With lv
        .view = lvwReport
        .Gridlines = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:="Transfer Instruction", Width:=144
    End With
End Sub

Public Sub ItemsToListView(ByVal lv As ListView)
    lv.ListItems.Clear
    
    Dim i As Long
    For i = 1 To Items.Count
        lv.ListItems.Add text:=Items(i).Name, key:="KEY" & i
    Next i
    
    If HasHistory Then
        If lv.ListItems.Count = 0 Then
            lv.ListItems.Add text:="(No items found)" ' TODO Const
        End If
    Else
        lv.ListItems.Add text:="(No history found)" ' TODO Const
    End If
End Sub

Public Sub TrySelect(ByVal key As String)
    Dim i As Long
    Debug.Assert Len(key) > 3
    i = val(Right$(key, Len(key) - 3))
    Debug.Assert i <= Items.Count
    Set SelectedInstruction = Items(i)
End Sub
