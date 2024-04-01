Attribute VB_Name = "TransferEnums"
'@Folder "MVVM.Resources.Constants"
Option Explicit

Public Enum TtViewResult
    vrCancel = 0
    vrOK
    vrStart
    vrBack
    vrNext
    vrFinish
End Enum

Public Enum TtDirection
    ttInvalidDirection = 0
    ttSource
    ttDestination
End Enum

Public Enum TtNode
    ttInvalidNode = 0
    ttApplication
    ttWorkbook
    ttWorksheet
    ttListObject
End Enum

Public Enum TtChangeType
    ttInvalidType = 0
    ttBlankUnchanged ' 0->0
    ttValueReplacesBlank ' A->0
    ttBlankReplacesValue ' 0->A
    ttValueUnchanged ' A->A
    ttValueChanged ' A->B
End Enum

Public Enum TtDeltaType
    ttInvalidMember = 0
    ttKeyMember
    ttField
    ttDelta
End Enum
