VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.ColumnPair"
Option Explicit
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    
    Items As Collection
    Selected As ColumnPair
End Type
Private This As TState

'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.Items.[_NewEnum]
End Property

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Public Property Get Selected() As ColumnPair
    Set Selected = This.Selected
End Property

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
    Set This.Items = New Collection
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

Public Function GetArray(ByVal Direction As TtDirection) As Variant
    If This.Items.Count = 0 Then
        Debug.Print "XYZ FOOBAR"
        GetArray = Array()
        Exit Function
    End If
    
    Dim Result() As Variant
    ReDim Result(0 To This.Items.Count - 1)
    
    Dim i As Long
    For i = 0 To UBound(Result)
        If Direction = ttSource Then
            Result(i) = This.Items.Item(i + 1).Source
        Else
            Result(i) = This.Items.Item(i + 1).Destination
        End If
    Next i
    
    GetArray = Result
End Function

Public Sub Add(ByVal Source As String, ByVal Destination As String)
    If Exists(Source, Destination) Then Exit Sub
    
    Dim ColumnPair As ColumnPair
    Set ColumnPair = New ColumnPair
    With ColumnPair
        .Source = Source
        .Destination = Destination
        .Key = Source & Chr$(31) & Destination
    End With
    
    This.Items.Add Key:=ColumnPair.Key, Item:=ColumnPair
End Sub

Public Sub Remove(ByVal Source As String, ByVal Destination As String)
    Dim i As Long
    For i = 1 To This.Items.Count
        Dim ColumnPair As ColumnPair
        Set ColumnPair = This.Items.Item(i)
        If ColumnPair.Source = Source And ColumnPair.Destination = Destination Then
            This.Items.Remove i
            Exit Sub
        End If
    Next i
End Sub

Public Sub Clear()
    CollectionClear This.Items
End Sub

Public Function Exists(ByVal Source As String, ByVal Destination As String) As Boolean
    Dim ColumnPair As ColumnPair
    Exists = TryGetColumnPair(Source, Destination, ColumnPair)
End Function

Public Sub TrySelect(ByVal Key As String)
    If TryGetColumnPair(Split(Key, Chr$(31))(0), Split(Key, Chr$(31))(1), This.Selected) Then
        OnPropertyChanged "Selected"
    End If
End Sub

Private Function TryGetColumnPair(ByVal Source As String, ByVal Destination As String, ByRef OutColumnPair As ColumnPair) As Boolean
    Dim ThisColumnPair As ColumnPair
    For Each ThisColumnPair In This.Items
        If ThisColumnPair.Source = Source And ThisColumnPair.Destination = Destination Then
            Set OutColumnPair = ThisColumnPair
            TryGetColumnPair = True
            Exit Function
        End If
    Next ThisColumnPair
End Function
