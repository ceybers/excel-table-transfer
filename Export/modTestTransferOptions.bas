Attribute VB_Name = "modTestTransferOptions"
'@Folder("TransferOptions")
Option Explicit
Option Private Module

Public Enum TransferOptionsEnum
    Invalid = 2 ^ 0
    
    ClearDestinationFirst = 2 ^ 1 ' 2
    TransferBlanks = 2 ^ 2 '4
    ReplaceEmptyOnly = 2 ^ 3 ' 8
    
    SourceFilteredOnly = 2 ^ 4 ' 16
    DestinationFilteredOnly = 2 ^ 5 ' 32
End Enum

Public Sub TestTransferOptions()
    Dim vm As IViewModel
    Set vm = Nothing
    
    Dim view As IView
    Set view = New TransferOptionsView
    
    Dim vview As TransferOptionsView
    Set vview = view
    
    If view.ShowDialog(vm) Then
        Debug.Print "Result = " & vview.flags
        
        MsgBox "Do we clear dest first? " & HasFlag(vview.flags, TransferOptionsEnum.ClearDestinationFirst)
    Else
        Debug.Print "Cancelled"
    End If
End Sub

Public Sub TestTransferOptions2()
    Dim flags As Integer
    flags = TransferOptionsEnum.ClearDestinationFirst
    
    PrintFlags flags
    
    flags = AddFlag(flags, TransferBlanks)
    flags = RemoveFlag(flags, Invalid)
    flags = RemoveFlag(flags, ClearDestinationFirst)
     
    PrintFlags flags
End Sub

Public Function AddFlag(ByVal flags As Integer, ByVal flag As Integer) As Integer
    If Not HasFlag(flags, flag) Then
        AddFlag = flags + flag
    Else
        AddFlag = flags
    End If
End Function

Public Function RemoveFlag(ByVal flags As Integer, ByVal flag As Integer) As Integer
    If HasFlag(flags, flag) Then
        RemoveFlag = flags - flag
    Else
        RemoveFlag = flags
    End If
End Function

Public Function HasFlag(ByVal flags As Integer, ByVal flag As Integer) As Boolean
    HasFlag = (flags And flag) = flag
End Function

Public Function SetFlag(ByVal flags As Integer, ByVal flag As Integer, ByVal checked As Boolean) As Integer
    If checked Then
        SetFlag = AddFlag(flags, flag)
    Else
        SetFlag = RemoveFlag(flags, flag)
    End If
End Function

Private Sub PrintFlags(ByVal flags As Integer)
    Debug.Print "TEST"
    Debug.Print "===="
    Debug.Print "Has Invalid: " & HasFlag(flags, Invalid)
    Debug.Print "Has ClearDestinationFirst: " & HasFlag(flags, ClearDestinationFirst)
    Debug.Print "Has TransferBlanks: " & HasFlag(flags, TransferBlanks)
    Debug.Print "Has ReplaceEmptyOnly: " & HasFlag(flags, ReplaceEmptyOnly)
    Debug.Print vbNullString
End Sub
