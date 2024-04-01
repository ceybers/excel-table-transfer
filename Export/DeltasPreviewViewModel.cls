VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeltasPreviewViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ViewModels"
Option Explicit
Implements INotifyPropertyChanged

Private Const MAX_LIMIT_DELTAS As Long = 1000000

Private Type TState
    Notifier As INotifyPropertyChanged
    
    TransferDeltas As TransferDeltas
    Source As ListObject ' Reference so we can get NumberFormats per ListColumn
    Destination As ListObject ' Reference so we can get NumberFormats per ListColumn
End Type
Private This As TState

Public Property Get Keys() As Collection
    Set Keys = This.TransferDeltas.Keys
End Property

Public Property Get Fields() As Collection
    Set Fields = This.TransferDeltas.Fields
End Property

Public Property Get Deltas() As TransferDeltas
    Set Deltas = This.TransferDeltas
End Property

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
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

Public Sub Load(ByVal TransferDeltas As TransferDeltas, ByVal Source As ListObject, ByVal Destination As ListObject)
    Set This.TransferDeltas = TransferDeltas
    Set This.Source = Source
    Set This.Destination = Destination
    
    This.TransferDeltas.SetNumberFormats GetNumberFormats(ttSource), GetNumberFormats(ttDestination)
End Sub

Private Function GetNumberFormats(ByVal Direction As TtDirection) As Object
    Dim Result As Object
    Set Result = CreateObject("Scripting.Dictionary")
    
    Dim ListColumns As ListColumns
    If Direction = ttSource Then
        Set ListColumns = This.Source.ListColumns
    Else
        Set ListColumns = This.Destination.ListColumns
    End If
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListColumns
        Result(ListColumn.Name) = ListColumn.DataBodyRange.Cells.Item(1, 1).NumberFormat
        If Result(ListColumn.Name) = "General" Then Result(ListColumn.Name) = "@"
    Next ListColumn
    
    Set GetNumberFormats = Result
End Function

Public Sub TrySelectKey(ByVal Key As String)
    This.TransferDeltas.KeyFilter = Key
End Sub

Public Sub TrySelectField(ByVal Field As String)
    This.TransferDeltas.FieldFilter = Field
End Sub

Public Function CanShowAll() As Boolean
    CanShowAll = This.TransferDeltas.Items.Count > This.TransferDeltas.Limit
End Function

Public Function CanFinish() As Boolean
    CanFinish = This.TransferDeltas.Items.Count > 0
End Function

Public Sub DoShowAll()
    This.TransferDeltas.Limit = MAX_LIMIT_DELTAS
    OnPropertyChanged "Deltas"
End Sub

