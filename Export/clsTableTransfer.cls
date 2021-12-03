VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents tableBrowser As frmTableBrowser
Attribute tableBrowser.VB_VarHelpID = -1
Private WithEvents keyMapper As frmKeyMapper
Attribute keyMapper.VB_VarHelpID = -1
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
    
    Set tableBrowser = New frmTableBrowser
    Set keyMapper = New frmKeyMapper
    Set valueMapper = New frmValueMapper
End Sub

Public Sub TransferFrom(lo As ListObject)
    Set this.Source = lo
    GetDestinationTable
End Sub

Private Sub GetDestinationTable()
    tableBrowser.Show
End Sub

Private Sub GetKeyMapping()
    Call keyMapper.SetTables(this.Source, this.Destination)
    keyMapper.Show
End Sub

Private Sub GetValueMapping()
    'MsgBox ("Now map the columns")
    
    'Debug.Print "srcKeyCol = " & this.SourceKeyColumn.Name
    'Debug.Print "dstKeyCol = " & this.DestinationKeyColumn.Name
    
'    Call frmColumnMapper.PopulateColumnMapper(GetSelectedTable(), lo)
'    frmColumnMapper.Show
    With this
        Call valueMapper.PopulateColumnMapper(.Source, .SourceKeyColumn, .Destination, .DestinationKeyColumn)
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

Private Sub keyMapper_Completed(lhs As ListColumn, rhs As ListColumn)
    Set this.SourceKeyColumn = lhs
    Set this.DestinationKeyColumn = rhs
    GetValueMapping
End Sub

Private Sub tableBrowser_TableSelected(ByVal lo As ListObject)
    Set this.Destination = lo
    GetKeyMapping
End Sub

Private Sub valueMapper_Cancelled()
    Debug.Print "Value mapper cancelled"
End Sub

Private Sub valueMapper_Completed(result As Variant)
    this.ValueMapping = result
    this.ClearDestinationColumns = valueMapper.chkClearDstCols.Value
    this.PasteIntoBlankCellsOnly = valueMapper.chkPasteIntoBlanks.Value
    this.CopyNonBlankCellsOnly = valueMapper.chkCopyNonBlankCellsOnly.Value
    DoTransfer
End Sub
