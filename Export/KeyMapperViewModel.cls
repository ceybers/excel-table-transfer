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
Implements INotifyPropertyChanged
Implements IHandlePropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    
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
    
    Set This.Notifier = New PropertyChangeNotifier
    RegisterWithChild This.Source
    RegisterWithChild This.Destination
End Sub

Private Sub RegisterWithChild(ByVal Model As INotifyPropertyChanged)
    Model.RegisterHandler Me
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler Handler
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    INotifyPropertyChanged_OnPropertyChanged Me, PropertyName
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If PropertyName = "Caption" Then
        TryEvaluateMatch
    ElseIf Source Is This.Source Then
        OnPropertyChanged "Source"
    ElseIf Source Is This.Destination Then
        OnPropertyChanged "Destination"
    End If
End Sub

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
    
    OnPropertyChanged "MatchQualityCaption"
    OnPropertyChanged "CanNext"
End Sub

Public Property Get CanNext() As Boolean
    If This.MatchQuality Is Nothing Then Exit Property
    If This.MatchQuality.Intersection.Count = 0 Then Exit Property
    
    CanNext = True
End Property

Public Property Get MatchQualityCaption()
    If This.MatchQuality Is Nothing Then
        MatchQualityCaption = NO_TWO_COLUMNS_SELECTED
    Else
        MatchQualityCaption = GetKeyColumnComparerString(This.MatchQuality)
    End If
End Property

Private Function GetKeyColumnComparerString(ByVal Comparer As KeyColumnComparer) As String
    Dim Intersection As Long
    Intersection = Comparer.Intersection.Count
    
    Dim Total As Long
    Total = Comparer.LeftOnly.Count + Comparer.Intersection.Count + Comparer.RightOnly.Count
    
    If Total = 0 Then
        GetKeyColumnComparerString = MSG_ZERO_KEYS
    Else
        GetKeyColumnComparerString = Format$(Intersection / Total, "0%") & " (" & CStr(Intersection) & _
            "/" & CStr(Total) & " keys intersect)"
    End If
End Function
