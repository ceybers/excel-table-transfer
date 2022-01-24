VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VBAProject"
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents keyMapper As frmKeyMapper
Attribute keyMapper.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents valueMapper As frmValueMapper
Attribute valueMapper.VB_VarHelpID = -1

Private Type tTableTransfer
    Source As ListObject
    Destination As ListObject
    SourceKeyColumn As ListColumn
    DestinationKeyColumn As ListColumn
    ValueMapping As Variant
    ClearDestinationColumns As Boolean
    PasteIntoBlankCellsOnly As Boolean
    CopyNonBlankCellsOnly As Boolean
End Type

Private this As tTableTransfer

Private Sub Class_Initialize()
    'TODO Use active selection picker instead
    'Set this.Source = ThisWorkbook.Worksheets("Sheet1").listobjects("Table1")
    
    Set keyMapper = New frmKeyMapper
    Set valueMapper = New frmValueMapper
End Sub

Public Sub TransferFrom(ByRef lo As ListObject)
    Set this.Source = lo

    Dim vm As clsSelectTableViewModel
    Set vm = New clsSelectTableViewModel
    Set vm.ActiveTable = lo
    
    Dim frm As IView
    Set frm = New frmSelectTable
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set this.Destination = vm.SelectedTable
            GetKeyMapping
        End If
    End If
End Sub

Private Sub GetKeyMapping()
    keyMapper.SetTables this.Source, this.Destination
    keyMapper.Show
End Sub

Private Sub GetValueMapping()
    With this
        valueMapper.PopulateColumnMapper .Source, .SourceKeyColumn, .Destination, .DestinationKeyColumn
    End With
    valueMapper.Show
End Sub

Private Sub DoTransfer()
    Dim ti As TransferInstruction
    With ti
        Set .Source = this.Source
        Set .SourceKey = this.SourceKeyColumn
        Set .Destination = this.Destination
        Set .DestinationKey = this.DestinationKeyColumn
        .Columns = this.ValueMapping
        .ClearDestinationColumns = this.ClearDestinationColumns
        .PasteIntoBlankCellsOnly = this.PasteIntoBlankCellsOnly
        .CopyNonBlankCellsOnly = this.CopyNonBlankCellsOnly
    End With
    RunTransferInstruction ti
End Sub

Private Sub keyMapper_Completed(LHS As ListColumn, RHS As ListColumn)
    Set this.SourceKeyColumn = LHS
    Set this.DestinationKeyColumn = RHS
    GetValueMapping
End Sub

Private Sub valueMapper_Cancelled()
    Debug.Print "Value mapper cancelled"
End Sub

Private Sub valueMapper_Completed(result As Variant)
    this.ValueMapping = result
    this.ClearDestinationColumns = valueMapper.chkClearDstCols.value
    this.PasteIntoBlankCellsOnly = valueMapper.chkPasteIntoBlanks.value
    this.CopyNonBlankCellsOnly = valueMapper.chkCopyNonBlankCellsOnly.value
    DoTransfer
End Sub
