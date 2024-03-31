Attribute VB_Name = "ChangeType"
'@Folder "MVVM.Model.TransferInstruction"
Option Explicit

Public Function GetChangeType(ByVal LHS As Variant, ByVal RHS As Variant) As ChangeType2
    Dim LHSVarType As Long
    LHSVarType = GetVarTypeMod(LHS)
    
    Dim RHSVarType As Long
    RHSVarType = GetVarTypeMod(RHS)
    
    If LHSVarType = 10 Or RHSVarType = 10 Then GetChangeType = ttInvalidType
    
    If LHSVarType = 0 And RHSVarType = 0 Then GetChangeType = ttBlankUnchanged
    If LHSVarType = 0 And RHSVarType = 8 Then GetChangeType = ttBlankReplacesValue
    If LHSVarType = 8 And RHSVarType = 0 Then GetChangeType = ttValueReplacesBlank
    
    If LHSVarType = 8 And RHSVarType = 8 Then
        If LHS = RHS Then
            GetChangeType = ttValueUnchanged
        Else
            GetChangeType = ttValueChanged
        End If
    End If
End Function

Public Function ChangeTypeToString(ByVal Value As ChangeType2) As String
    Select Case Value
        Case ttInvalidType
            ChangeTypeToString = "Invalid"
        Case ttBlankUnchanged
            ChangeTypeToString = "BlankUnchanged"
        Case ttValueReplacesBlank
            ChangeTypeToString = "ValueReplacesBlank"
        Case ttBlankReplacesValue
            ChangeTypeToString = "BlankReplacesValue"
        Case ttValueUnchanged
            ChangeTypeToString = "ValueUnchanged"
        Case ttValueChanged
            ChangeTypeToString = "ValueChanged"
    End Select
End Function

Public Function GetVarTypeMod(ByVal Value As Variant) As Long
    Dim Result As Long
    Result = VarType(Value)
    
    Select Case Result
        Case 8
            If Len(Value) = 0 Then Result = 0
        Case 5
            Result = 8
        Case 11
            Result = 8
    End Select
    
    GetVarTypeMod = Result
End Function

