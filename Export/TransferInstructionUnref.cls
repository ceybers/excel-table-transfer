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
    SourceFilename As String
    DestinationFilename As String
    Source As String
    Destination As String
    SourceKey As String
    DestinationKey As String
    ValuePairs As Variant
    Flags As Integer
End Type

Private this As TState

Public Property Get SourceFilename() As String
    SourceFilename = this.SourceFilename
End Property

Public Property Let SourceFilename(ByVal RHS As String)
    this.SourceFilename = RHS
End Property

Public Property Get DestinationFilename() As String
    DestinationFilename = this.DestinationFilename
End Property

Public Property Let DestinationFilename(ByVal RHS As String)
    this.DestinationFilename = RHS
End Property

Public Property Get Source() As String
    Source = this.Source
End Property

Public Property Let Source(ByVal RHS As String)
    this.Source = RHS
End Property

Public Property Get Destination() As String
    Destination = this.Destination
End Property

Public Property Let Destination(ByVal RHS As String)
    this.Destination = RHS
End Property

Public Property Get SourceKey() As String
    SourceKey = this.SourceKey
End Property

Public Property Let SourceKey(ByVal RHS As String)
    this.SourceKey = RHS
End Property

Public Property Get DestinationKey() As String
    DestinationKey = this.DestinationKey
End Property

Public Property Let DestinationKey(ByVal RHS As String)
    this.DestinationKey = RHS
End Property

Public Property Get ValuePairs() As Variant
    Let ValuePairs = this.ValuePairs
End Property

Public Property Let ValuePairs(ByVal RHS As Variant)
    Let this.ValuePairs = RHS
End Property

Public Property Get Flags() As Integer
    Flags = this.Flags
End Property

Public Property Let Flags(ByVal RHS As Integer)
    this.Flags = RHS
End Property

' TODO Invert this so that TI can try and load from a TIUnref
' That way we can try and use the same ValuePair mapping on different tables
Public Function AsReferenced() As TransferInstruction
    Dim ti As TransferInstruction
    Set ti = New TransferInstruction
    
    ti.Flags = this.Flags
    
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

   
