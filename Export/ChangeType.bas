Attribute VB_Name = "ChangeType"
'@Folder "Model2.TransferInstruction2"
Option Explicit


Public Enum ChangeTypeEnum
    Invalid
    BlankUnchanged ' 0->0
    ValueReplacesBlank ' A->0
    BlankReplacesValue ' 0->A
    ValueUnchanged ' A->A
    ValueChanged ' A->B
End Enum

Public Function GetChangeType(ByVal LHS As Variant, ByVal RHS As Variant) As ChangeTypeEnum
    Dim LHSVarType As Long
    LHSVarType = GetVarTypeMod(LHS)
    
    Dim RHSVarType As Long
    RHSVarType = GetVarTypeMod(RHS)
    
    If LHSVarType = 10 Or RHSVarType = 10 Then GetChangeType = Invalid
    
    If LHSVarType = 0 And RHSVarType = 0 Then GetChangeType = BlankUnchanged
    If LHSVarType = 0 And RHSVarType = 8 Then GetChangeType = BlankReplacesValue
    If LHSVarType = 8 And RHSVarType = 0 Then GetChangeType = ValueReplacesBlank
    
    If LHSVarType = 8 And RHSVarType = 8 Then
        If LHS = RHS Then
            GetChangeType = ValueUnchanged
        Else
            GetChangeType = ValueChanged
        End If
    End If
End Function

Public Function ChangeTypeToString(ByVal Value As ChangeTypeEnum) As String
    Select Case Value
        Case ChangeTypeEnum.Invalid
            ChangeTypeToString = "Invalid"
        Case ChangeTypeEnum.BlankUnchanged
            ChangeTypeToString = "BlankUnchanged"
        Case ChangeTypeEnum.ValueReplacesBlank
            ChangeTypeToString = "ValueReplacesBlank"
        Case ChangeTypeEnum.BlankReplacesValue
            ChangeTypeToString = "BlankReplacesValue"
        Case ChangeTypeEnum.ValueUnchanged
            ChangeTypeToString = "ValueUnchanged"
        Case ChangeTypeEnum.ValueChanged
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

