Attribute VB_Name = "TransferCellDelta"
'@Folder "Model.TransferInstruction2"
Option Explicit

Public Function GetTransferCellDelta(ByVal KeyIndex As Long, ByVal ValueIndex As Long, _
    ByVal Key As String, ByVal SourceColumnName As String, ByVal DestinationColumnName As String, _
    ByVal SourceValue As Variant, ByVal DestinationValue As Variant) As Variant
    
    'Debug.Print " CHG: '"; Key; "' ("; CStr(Index); ") x ";
    'Debug.Print "('"; SourceColumnName; "', '"; DestinationColumnName; "') => ";
    'Debug.Print "'"; SourceValue; "' --> '"; DestinationValue; "' ";
    'Debug.Print "("; CStr(VarType(SourceValue)); ","; CStr(VarType(DestinationValue)); ", ";
    'Debug.Print " LHS LEN = "; Len(CStr(SourceValue));
    'Debug.Print ChangeTypeToString(ChangeType.GetChangeType(SourceValue, DestinationValue)); ")"
    
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
    Result(2) = CellDelta(2)
    Result(3) = "' x ('"
    Result(4) = CellDelta(3)
    Result(5) = "','"
    Result(6) = CellDelta(4)
    Result(7) = "') = {("
    Result(8) = CellDelta(0)
    Result(9) = ","
    Result(10) = CellDelta(1)
    Result(11) = "); "
    Result(12) = CStr(CellDelta(5))
    Result(13) = "' -> '"
    Result(14) = CStr(CellDelta(6))
    Result(15) = "'} ["
    Result(16) = CStr(VarType(CellDelta(5)))
    Result(17) = ","
    Result(18) = CStr(VarType(CellDelta(6)))
    Result(19) = ","
    Result(20) = ChangeType.ChangeTypeToString(CellDelta(7))
    Result(21) = "]"
    
    ToString = Join(Result, vbNullString)
End Function


