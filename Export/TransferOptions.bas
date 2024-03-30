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
