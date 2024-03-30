Attribute VB_Name = "TransferCellDelta"
'@Folder "Model.TransferInstruction2"
Option Explicit

Public Enum DeltaIndex
    tdRow = 0
    tdCol
    tdKey
    tdFieldSrc
    tdFieldDst
    tdValueBefore
    tdValueAfter
    'tdVarTypeBefore
    'tdVarTypeAfter
    tdChangeType
End Enum

Public Function GetTransferCellDelta(ByVal KeyIndex As Long, ByVal ValueIndex As Long, _
    ByVal Key As String, ByVal SourceColumnName As String, ByVal DestinationColumnName As String, _
    ByVal SourceValue As Variant, ByVal DestinationValue As Variant) As Variant
    
    Dim Result As Variant
    Result = Array(KeyIndex, ValueIndex, _
        Key, SourceColumnName, DestinationColumnName, _
        SourceValue, DestinationValue, _
        ChangeType.GetChangeType(SourceValue, DestinationValue))
        
    GetTransferCellDelta = Result
End Function

Public Function ToString(ByVal CellDelta As Variant) As String
    Debug.Assert Not IsEmpty(CellDelta)
    
    Dim Result(1 To 21) As String
    Result(1) = "CHG: '"
    Result(2) = CellDelta(tdKey)
    Result(3) = "' x ('"
    Result(4) = CellDelta(tdFieldSrc)
    Result(5) = "','"
    Result(6) = CellDelta(tdFieldDst)
    Result(7) = "') = {("
    Result(8) = CellDelta(tdRow)
    Result(9) = ","
    Result(10) = CellDelta(tdCol)
    Result(11) = "); "
    Result(12) = CStr(CellDelta(tdValueBefore))
    Result(13) = "' -> '"
    Result(14) = CStr(CellDelta(tdValueAfter))
    Result(15) = "'} ["
    'Result(16) = CStr(VarType(CellDelta(tdVarTypeBefore)))
    'Result(17) = ","
    'Result(18) = CStr(VarType(CellDelta(tdVarTypeAfter)))
    'Result(19) = ","
    Result(20) = ChangeType.ChangeTypeToString(CellDelta(tdChangeType))
    Result(21) = "]"
    
    ToString = Join(Result, vbNullString)
End Function


