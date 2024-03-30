VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceOrDestinationViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ViewModels"
Option Explicit
Implements IViewModel

Private Type TSourceOrDestinationViewModel
    IsSource As Boolean
    IsDestination As Boolean
    ListObject As ListObject
End Type

Private This As TSourceOrDestinationViewModel

Public Property Get IsSource() As Boolean
    IsSource = This.IsSource
End Property

Public Property Let IsSource(ByVal RHS As Boolean)
    This.IsSource = RHS
End Property

Public Property Get IsDestination() As Boolean
    IsDestination = This.IsDestination
End Property

Public Property Let IsDestination(ByVal RHS As Boolean)
    This.IsDestination = RHS
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Set ListObject(ByVal RHS As ListObject)
    Set This.ListObject = RHS
End Property

Private Function IViewModel_IsValid() As Boolean
    If IsSource = False And IsDestination = False Then Exit Function
    IViewModel_IsValid = True
End Function


Public Function CanSelect() As Boolean
    Dim AllTables As Collection
    Set AllTables = ListObjectHelpers.GetAllTablesInApplication
    If AllTables Is Nothing Then Exit Function
    If AllTables.Count = 0 Then Exit Function
    
    CanSelect = True
End Function

