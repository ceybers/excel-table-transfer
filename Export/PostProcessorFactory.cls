VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostProcessorFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Model2.PostProcessors"
Option Explicit

Public Property Get HighlightChanges() As ITransferPostProcesser
    Set HighlightChanges = New HighlightChanges
End Property

Public Property Get RemoveHighlighting() As ITransferPostProcesser
    Set RemoveHighlighting = New RemoveHighlighting
End Property

Public Property Get TransferColumnWidths() As ITransferPostProcesser
    Set TransferColumnWidths = New TransferColumnWidths
End Property

Public Property Get TransferNumberFormat() As ITransferPostProcesser
    Set TransferNumberFormat = New TransferNumberFormat
End Property
