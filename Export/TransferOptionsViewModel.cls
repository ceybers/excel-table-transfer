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

Private Function IViewModel_IsValid() As Boolean
    Err.Raise 5                                  'TODO implement interface member
End Function
