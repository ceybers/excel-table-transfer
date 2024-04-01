Attribute VB_Name = "ChangeType"
'@Folder "MVVM.Model.TransferInstruction"
Option Explicit

Public Function GetChangeType(ByVal LHS As Variant, ByVal RHS As Variant) As TtChangeType
    Dim LHSVarType As Long
    LHSVarType = GetVarTypeMod(LHS)
    
    Dim RHSVarType As Long
    RHSVarType = GetVarTypeMod(RHS)
    
    If LHSVarType = vbError Or RHSVarType = vbError Then GetChangeType = ttInvalidType
    
    If LHSVarType = vbEmpty And RHSVarType = vbEmpty Then GetChangeType = ttBlankUnchanged
    If LHSVarType = vbEmpty And RHSVarType = vbString Then GetChangeType = ttBlankReplacesValue
    If LHSVarType = vbString And RHSVarType = vbEmpty Then GetChangeType = ttValueReplacesBlank
    
    If LHSVarType = vbString And RHSVarType = vbString Then
        If LHS = RHS Then
            GetChangeType = ttValueUnchanged
        Else
            GetChangeType = ttValueChanged
        End If
    End If
End Function

Public Function ChangeTypeToString(ByVal Value As TtChangeType) As String
    Select Case Value
        Case ttInvalidType
            ChangeTypeToString = CHANGE_TYPE_INVALID
        Case ttBlankUnchanged
            ChangeTypeToString = CHANGE_TYPE_BLANK_UNCHANGED
        Case ttValueReplacesBlank
            ChangeTypeToString = CHANGE_TYPE_VALUE_REPLACES_BLANK
        Case ttBlankReplacesValue
            ChangeTypeToString = CHANGE_TYPE_BLANK_REPLACES_VALUE
        Case ttValueUnchanged
            ChangeTypeToString = CHANGE_TYPE_VALUE_UNCHANGED
        Case ttValueChanged
            ChangeTypeToString = CHANGE_TYPE_VALUE_CHANGED
    End Select
End Function

Public Function GetVarTypeMod(ByVal Value As Variant) As Long
    Dim Result As Long
    Result = VarType(Value)
    
    Select Case Result
        Case vbString
            If LenB(Value) = 0 Then Result = 0
        Case vbDouble
            Result = vbString
        Case vbBoolean
            Result = vbString
    End Select
    
    GetVarTypeMod = Result
End Function

