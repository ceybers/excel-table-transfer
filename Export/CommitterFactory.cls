VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommitterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Model2.Commiter"
Option Explicit

Public Property Get FullColumn() As ICommitStrategy
    Set FullColumn = New CommitFullColumn
End Property
