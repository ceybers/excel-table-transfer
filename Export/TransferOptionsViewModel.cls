VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferOptionsViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ViewModels"
Option Explicit
Implements IViewModel

Private Type TViewModel
    ' TODO Reset to before changes
    ' i.e. StartEditing, StopEditing, CancelEditing
    ' OldFlags as integer
    Flags As Long
End Type

Private This As TViewModel

Public Property Get Flags() As Long
    Flags = This.Flags
End Property

Public Property Let Flags(ByVal RHS As Long)
    This.Flags = RHS
End Property

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

