VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceOrDestinationViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "SourceOrDestination"
Option Explicit
Implements IViewModel

Private Type TSourceOrDestinationViewModel
    IsSource As Boolean
    IsDestination As Boolean
    ListObject As ListObject
End Type

Private this As TSourceOrDestinationViewModel

Public Property Get IsSource() As Boolean
    IsSource = this.IsSource
End Property

Public Property Let IsSource(ByVal RHS As Boolean)
    this.IsSource = RHS
End Property

Public Property Get IsDestination() As Boolean
    IsDestination = this.IsDestination
End Property

Public Property Let IsDestination(ByVal RHS As Boolean)
    this.IsDestination = RHS
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = this.ListObject
End Property

Public Property Set ListObject(ByVal RHS As ListObject)
    Set this.ListObject = RHS
End Property

Private Function IViewModel_IsValid() As Boolean
    If IsSource = False And IsDestination = False Then Exit Function
    IViewModel_IsValid = True
End Function
