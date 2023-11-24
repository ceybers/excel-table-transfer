Attribute VB_Name = "modTestTransferOptions"
'@Folder("TransferOptions")
Option Explicit
Option Private Module

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

Public Sub Test()
    Dim vm As TransferOptionsViewModel
    Set vm = New TransferOptionsViewModel
    vm.Flags = ClearDestinationFirst + ReplaceEmptyOnly + DestinationFilteredOnly
    
    Dim view As IView
    Set view = New TransferOptionsView
    
    If view.ShowDialog(vm) Then
        PrintFlags vm.Flags
    Else
        Debug.Print "Cancelled"
    End If
End Sub

Public Sub TestTransferOptions2()
    Dim Flags As Integer
    Flags = TransferOptionsEnum.ClearDestinationFirst
    
    PrintFlags Flags
    
    Flags = AddFlag(Flags, TransferBlanks)
    Flags = RemoveFlag(Flags, Invalid)
    Flags = RemoveFlag(Flags, ClearDestinationFirst)
     
    PrintFlags Flags
End Sub

Public Function AddFlag(ByVal Flags As Integer, ByVal flag As TransferOptionsEnum) As Integer
    If Not HasFlag(Flags, flag) Then
        AddFlag = Flags + flag
    Else
        AddFlag = Flags
    End If
End Function

Public Function RemoveFlag(ByVal Flags As Integer, ByVal flag As TransferOptionsEnum) As Integer
    If HasFlag(Flags, flag) Then
        RemoveFlag = Flags - flag
    Else
        RemoveFlag = Flags
    End If
End Function

Public Function HasFlag(ByVal Flags As Integer, ByVal flag As TransferOptionsEnum) As Boolean
    HasFlag = (Flags And flag) = flag
End Function

Public Function SetFlag(ByVal Flags As Integer, ByVal flag As TransferOptionsEnum, ByVal checked As Boolean) As Integer
    If checked Then
        SetFlag = AddFlag(Flags, flag)
    Else
        SetFlag = RemoveFlag(Flags, flag)
    End If
End Function

Private Sub PrintFlags(ByVal Flags As Integer)
    Debug.Print "TEST"
    Debug.Print "===="
    Debug.Print "Has Invalid: " & HasFlag(Flags, Invalid)
    Debug.Print "Has ClearDestinationFirst: " & HasFlag(Flags, ClearDestinationFirst)
    Debug.Print "Has TransferBlanks: " & HasFlag(Flags, TransferBlanks)
    Debug.Print "Has ReplaceEmptyOnly: " & HasFlag(Flags, ReplaceEmptyOnly)
    Debug.Print "Has SourceFilteredOnly: " & HasFlag(Flags, SourceFilteredOnly)
    Debug.Print "Has DestinationFilteredOnly: " & HasFlag(Flags, DestinationFilteredOnly)
    Debug.Print "Has RemoveUnmapped: " & HasFlag(Flags, RemoveUnmapped)
    Debug.Print "Has AppendUnmapped: " & HasFlag(Flags, AppendUnmapped)
    Debug.Print vbNullString
End Sub

