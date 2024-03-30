Attribute VB_Name = "TransferOptions"
'@Folder "Model.TransferOptions"
Option Explicit

Public Enum TransferOptionsEnum
    Invalid = 2 ^ 0
    
    ClearDestinationFirst = 2 ^ 1                ' 2
    TransferBlanks = 2 ^ 2                       '4
    ReplaceEmptyOnly = 2 ^ 3                     ' 8
    
    SourceFilteredOnly = 2 ^ 4                   ' 16
    DestinationFilteredOnly = 2 ^ 5              ' 32
    
    RemoveUnmapped = 2 ^ 6                       ' 64
    AppendUnmapped = 2 ^ 7                       ' 128
    
    SaveToHistory = 2 ^ 8                        ' 256
    
    HighlightMapped = 2 ^ 9                      ' 512
End Enum

Public Function AddFlag(ByVal Flags As Long, ByVal Flag As TransferOptionsEnum) As Long
    If Not HasFlag(Flags, Flag) Then
        AddFlag = Flags + Flag
    Else
        AddFlag = Flags
    End If
End Function

Public Function RemoveFlag(ByVal Flags As Long, ByVal Flag As TransferOptionsEnum) As Long
    If HasFlag(Flags, Flag) Then
        RemoveFlag = Flags - Flag
    Else
        RemoveFlag = Flags
    End If
End Function

Public Function HasFlag(ByVal Flags As Long, ByVal Flag As TransferOptionsEnum) As Boolean
    HasFlag = (Flags And Flag) = Flag
End Function

Public Function SetFlag(ByVal Flags As Long, ByVal Flag As TransferOptionsEnum, ByVal Checked As Boolean) As Long
    If Checked Then
        SetFlag = AddFlag(Flags, Flag)
    Else
        SetFlag = RemoveFlag(Flags, Flag)
    End If
End Function
