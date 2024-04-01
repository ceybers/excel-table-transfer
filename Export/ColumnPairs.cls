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

Private Type TState
    Items As Collection
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

Private Sub Class_Initialize()
    Set This.Items = New Collection
End Sub

Public Function GetArray(ByVal Direction As TtDirection) As Variant
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
    End With
    
    This.Items.Add Item:=ColumnPair
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
