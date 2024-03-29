VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferInstructionUnref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule
'@Folder "Model.TransferHistory"
Option Explicit

Private Type TState
    SourceSheet As String
    DestinationSheet As String
    SourceFilename As String
    DestinationFilename As String
    Source As String
    Destination As String
    SourceKey As String
    DestinationKey As String
    ValuePairs As Variant
    Flags As Integer
    LastRun As Date
End Type

Private This As TState

Public Property Get LastRun() As Date
    LastRun = This.LastRun
End Property

Public Property Let LastRun(ByVal RHS As Date)
    This.LastRun = RHS
End Property

Public Property Get SourceSheet() As String
    SourceSheet = This.SourceSheet
End Property

Public Property Let SourceSheet(ByVal RHS As String)
    This.SourceSheet = RHS
End Property

Public Property Get DestinationSheet() As String
    DestinationSheet = This.DestinationSheet
End Property

Public Property Let DestinationSheet(ByVal RHS As String)
    This.DestinationSheet = RHS
End Property

Public Property Get SourceFilename() As String
    SourceFilename = This.SourceFilename
End Property

Public Property Let SourceFilename(ByVal RHS As String)
    This.SourceFilename = RHS
End Property

Public Property Get DestinationFilename() As String
    DestinationFilename = This.DestinationFilename
End Property

Public Property Let DestinationFilename(ByVal RHS As String)
    This.DestinationFilename = RHS
End Property

Public Property Get Source() As String
    Source = This.Source
End Property

Public Property Let Source(ByVal RHS As String)
    This.Source = RHS
End Property

Public Property Get Destination() As String
    Destination = This.Destination
End Property

Public Property Let Destination(ByVal RHS As String)
    This.Destination = RHS
End Property

Public Property Get SourceKey() As String
    SourceKey = This.SourceKey
End Property

Public Property Let SourceKey(ByVal RHS As String)
    This.SourceKey = RHS
End Property

Public Property Get DestinationKey() As String
    DestinationKey = This.DestinationKey
End Property

Public Property Let DestinationKey(ByVal RHS As String)
    This.DestinationKey = RHS
End Property

Public Property Get ValuePairs() As Variant
    Let ValuePairs = This.ValuePairs
End Property

Public Property Let ValuePairs(ByVal RHS As Variant)
    Let This.ValuePairs = RHS
End Property

Public Property Get Flags() As Integer
    Flags = This.Flags
End Property

Public Property Let Flags(ByVal RHS As Integer)
    This.Flags = RHS
End Property

' TODO Invert this so that TI can try and load from a TIUnref
' That way we can try and use the same ValuePair mapping on different tables
Public Function AsReferenced() As TransferInstruction
    Dim ti As TransferInstruction
    Set ti = New TransferInstruction
    
    ti.Flags = This.Flags
    'ti.LastRun = this.LastRun
    
    Dim wb As Workbook
    Dim wbLHS As Workbook
    Dim wbRHS As Workbook
    For Each wb In Workbooks
        If wb.Name = This.SourceFilename Then
            Set wbLHS = wb
        End If
        If wb.Name = This.DestinationFilename Then
            Set wbRHS = wb
        End If
    Next wb
    
    Dim lo As ListObject
    If Not wbLHS Is Nothing Then
        If TryGetListObjectFromWorkbook(wbLHS, This.Source, lo) Then
            Set ti.Source = lo
            On Error Resume Next
            Set ti.SourceKey = ti.Source.ListColumns.Item(This.SourceKey)
            On Error GoTo 0
        End If
    End If
    
    If Not wbRHS Is Nothing Then
        If TryGetListObjectFromWorkbook(wbRHS, This.Destination, lo) Then
            Set ti.Destination = lo
            On Error Resume Next
            Set ti.DestinationKey = ti.Destination.ListColumns.Item(This.DestinationKey)
            On Error GoTo 0
        End If
    End If
    
    Set AsReferenced = ti
End Function

Public Function LoadFromRange(ByVal rng As Range) As Boolean
    Dim c As Long
    c = 12 + rng.offset(10, 2).Value2
    
    Dim arr As Variant
    arr = rng.Resize(c, 4).Value2
    
    If arr(1, 1) <> "TRANSFER" Or arr(c, 1) <> "END" Then
        Exit Function
    End If
    
    This.SourceFilename = arr(5, 3)
    This.DestinationFilename = arr(5, 4)
    This.SourceSheet = arr(6, 3)
    This.DestinationSheet = arr(6, 4)
    This.Source = arr(8, 3)
    This.Destination = arr(8, 4)
    This.SourceKey = arr(9, 3)
    This.DestinationKey = arr(9, 4)
    This.Flags = arr(10, 3)
    
    Dim vpArr As Variant
    vpArr = rng.offset(11, 2).Resize(c - 12, 2).Value2
    This.ValuePairs = vpArr
    
    LoadFromRange = True
End Function

