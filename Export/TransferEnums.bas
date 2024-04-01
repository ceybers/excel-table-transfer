Attribute VB_Name = "TransferEnums"
'@Folder "MVVM.Resources.Constants"
Option Explicit

Public Enum ViewResult
    vrCancel = 0
    vrOK
    vrStart
    vrBack
    vrNext
    vrFinish
End Enum

Public Enum TransferDirection
    ttInvalidDirection = 0
    ttSource
    ttDestination
End Enum

Public Enum TransferNode
    ttInvalidNode = 0
    ttApplication
    ttWorkbook
    ttWorksheet
    ttListObject
End Enum

Public Enum ChangeType2
    ttInvalidType = 0
    ttBlankUnchanged ' 0->0
    ttValueReplacesBlank ' A->0
    ttBlankReplacesValue ' 0->A
    ttValueUnchanged ' A->A
    ttValueChanged ' A->B
End Enum

Public Enum MemberType
    ttInvalidMember = 0
    ttKeyMember
    ttField
    ttDelta
End Enum
