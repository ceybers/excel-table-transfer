VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferInstructionUnref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TransferHistory")
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

Private this As TState

Public Property Get LastRun() As Date
    LastRun = this.LastRun
End Property

Public Property Let LastRun(ByVal rhs As Date)
    this.LastRun = rhs
End Property

Public Property Get SourceSheet() As String
    SourceSheet = this.SourceSheet
End Property

Public Property Let SourceSheet(ByVal rhs As String)
    this.SourceSheet = rhs
End Property

Public Property Get DestinationSheet() As String
    DestinationSheet = this.DestinationSheet
End Property

Public Property Let DestinationSheet(ByVal rhs As String)
    this.DestinationSheet = rhs
End Property

Public Property Get SourceFilename() As String
    SourceFilename = this.SourceFilename
End Property

Public Property Let SourceFilename(ByVal rhs As String)
    this.SourceFilename = rhs
End Property

Public Property Get DestinationFilename() As String
    DestinationFilename = this.DestinationFilename
End Property

Public Property Let DestinationFilename(ByVal rhs As String)
    this.DestinationFilename = rhs
End Property

Public Property Get Source() As String
    Source = this.Source
End Property

Public Property Let Source(ByVal rhs As String)
    this.Source = rhs
End Property

Public Property Get Destination() As String
    Destination = this.Destination
End Property

Public Property Let Destination(ByVal rhs As String)
    this.Destination = rhs
End Property

Public Property Get SourceKey() As String
    SourceKey = this.SourceKey
End Property

Public Property Let SourceKey(ByVal rhs As String)
    this.SourceKey = rhs
End Property

Public Property Get DestinationKey() As String
    DestinationKey = this.DestinationKey
End Property

Public Property Let DestinationKey(ByVal rhs As String)
    this.DestinationKey = rhs
End Property

Public Property Get ValuePairs() As Variant
    Let ValuePairs = this.ValuePairs
End Property

Public Property Let ValuePairs(ByVal rhs As Variant)
    Let this.ValuePairs = rhs
End Property

Public Property Get Flags() As Integer
    Flags = this.Flags
End Property

Public Property Let Flags(ByVal rhs As Integer)
    this.Flags = rhs
End Property

' TODO Invert this so that TI can try and load from a TIUnref
' That way we can try and use the same ValuePair mapping on different tables
Public Function AsReferenced() As TransferInstruction
    Dim ti As TransferInstruction
    Set ti = New TransferInstruction
    
    ti.Flags = this.Flags
    'ti.LastRun = this.LastRun
    
    Dim wb As Workbook
    Dim wbLHS As Workbook
    Dim wbRHS As Workbook
    For Each wb In Workbooks
        If wb.Name = this.SourceFilename Then
            Set wbLHS = wb
        End If
        If wb.Name = this.DestinationFilename Then
            Set wbRHS = wb
        End If
    Next wb
    
    Dim lo As ListObject
    If Not wbLHS Is Nothing Then
        If TryGetListObjectFromWorkbook(wbLHS, this.Source, lo) Then
            Set ti.Source = lo
            On Error Resume Next
            Set ti.SourceKey = ti.Source.ListColumns(this.SourceKey)
            On Error GoTo 0
        End If
    End If
    
    If Not wbRHS Is Nothing Then
        If TryGetListObjectFromWorkbook(wbRHS, this.Destination, lo) Then
            Set ti.Destination = lo
            On Error Resume Next
            Set ti.DestinationKey = ti.Destination.ListColumns(this.DestinationKey)
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
    
    this.SourceFilename = arr(5, 3)
    this.DestinationFilename = arr(5, 4)
    this.SourceSheet = arr(6, 3)
    this.DestinationSheet = arr(6, 4)
    this.Source = arr(8, 3)
    this.Destination = arr(8, 4)
    this.SourceKey = arr(9, 3)
    this.DestinationKey = arr(9, 4)
    this.Flags = arr(10, 3)
    
    Dim vpArr As Variant
    vpArr = rng.offset(11, 2).Resize(c - 12, 2).Value2
    this.ValuePairs = vpArr
    
    LoadFromRange = True
End Function
