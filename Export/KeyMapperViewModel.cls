VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ViewModels"
Option Explicit
Implements IViewModel

Private Type TState
    Source As KeyColumns
    Destination As KeyColumns
    MatchQuality As KeyColumnComparer
End Type

Private This As TState

Public Property Get Source() As KeyColumns
    Set Source = This.Source
End Property

Public Property Get Destination() As KeyColumns
    Set Destination = This.Destination
End Property

Public Property Get MatchQuality() As KeyColumnComparer
    Set MatchQuality = This.MatchQuality
End Property

Private Sub Class_Initialize()
    Set This.Source = New KeyColumns
    Set This.Destination = New KeyColumns
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Public Sub Load(ByVal Source As ListObject, ByVal Destination As ListObject)
    This.Source.Load Source
    This.Destination.Load Destination
    
End Sub

Public Sub TryEvaluateMatch()
    If This.Source.Selected Is Nothing Or This.Destination.Selected Is Nothing Then
        Set This.MatchQuality = Nothing
    Else
        Set This.MatchQuality = New KeyColumnComparer
        Set This.MatchQuality.LHS = This.Source.Selected
        Set This.MatchQuality.RHS = This.Destination.Selected
    End If
End Sub

Public Sub ShowMatchDetails()
    MsgBox "NYI"
End Sub

Public Sub ShowQuality(ByVal Direction As TransferDirection)
    MsgBox "NYI"
End Sub

Public Function CanNext() As Boolean
    If This.MatchQuality Is Nothing Then Exit Function
    If This.MatchQuality.Intersection.Count = 0 Then Exit Function
    
    CanNext = True
End Function
