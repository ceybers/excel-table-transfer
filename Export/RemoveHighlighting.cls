VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoveHighlighting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model2.PostProcessors"
Option Explicit
Implements ITransferPostProcesser

Private Sub ITransferPostProcesser_Execute(ByVal TransferInstruction As TransferInstruction2)
    Highlighting.RemoveExistingHighlighting TransferInstruction.Destination.Table.Parent
End Sub
