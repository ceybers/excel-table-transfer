VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferOptionsViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TransferOptions"
Option Explicit
Implements IViewModel

Private Type TViewModel
    ' TODO Reset to before changes
    ' i.e. StartEditing, StopEditing, CancelEditing
    ' OldFlags as integer
    Flags As Integer
End Type

Private this As TViewModel

Public Property Get Flags() As Integer
    Flags = this.Flags
End Property

Public Property Let Flags(ByVal rhs As Integer)
    this.Flags = rhs
End Property

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function
