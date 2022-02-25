VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("KeyMapper")
Option Explicit
Implements IViewModel

Public Event PropertyChanged(ByVal propertyName As String)
Public Event CheckCompleted()
Public Event PreviewChanged()
Public Event MatchChanged()

Private Type TViewModel
    LHSTable As ListObject
    RHSTable As ListObject
    LHSKeyColumn As ListColumn
    RHSKeyColumn As ListColumn
    LHSColumns As Collection ' Remove?
    RHSColumns As Collection ' Remove?
    IsDirty As Boolean
    IsDirtyQuality As Boolean
    IsDirtyMatch As Boolean
    '---
    QualityLHS As KeyColumn
    QualityRHS As KeyColumn
    MatchResults As KeyColumnComparer
End Type

Private this As TViewModel

Public Property Get IsDirty() As Boolean
    IsDirty = this.IsDirty
End Property

Public Property Get LHSTable() As ListObject
    Set LHSTable = this.LHSTable
End Property

Public Property Set LHSTable(ByVal RHS As ListObject)
    If this.LHSTable Is RHS Then Exit Property
    Set this.LHSTable = RHS
    this.IsDirty = True
    Debug.Assert Not RHS Is Nothing
    UpdateColumns this.LHSTable, this.LHSColumns
    RaiseEvent PropertyChanged("LHSTable")
    RaiseEvent PropertyChanged("LHSColumns")
End Property

Public Property Get RHSTable() As ListObject
    Set RHSTable = this.RHSTable
End Property

Public Property Set RHSTable(ByVal RHS As ListObject)
    If this.RHSTable Is RHS Then Exit Property
    Set this.RHSTable = RHS
    this.IsDirty = True
    UpdateColumns this.RHSTable, this.RHSColumns
    RaiseEvent PropertyChanged("RHSTable")
    RaiseEvent PropertyChanged("RHSColumns")
End Property

Public Property Get LHSColumns() As Collection
    Set LHSColumns = this.LHSColumns
End Property

Public Property Set LHSColumns(ByVal RHS As Collection)
    Set this.LHSColumns = RHS
End Property

Public Property Get RHSColumns() As Collection
    Set RHSColumns = this.RHSColumns
End Property

Public Property Set RHSColumns(ByVal RHS As Collection)
    Set this.RHSColumns = RHS
End Property

Public Property Get LHSKeyColumn() As ListColumn
    Set LHSKeyColumn = this.LHSKeyColumn
End Property

Public Property Set LHSKeyColumn(ByVal RHS As ListColumn)
    If Not this.LHSKeyColumn Is Nothing Then
        If this.LHSKeyColumn.Name = RHS.Name Then Exit Property
    End If
    this.IsDirtyQuality = True
    this.IsDirtyMatch = False
    this.IsDirty = True
    Set this.LHSKeyColumn = RHS
    RaiseEvent PropertyChanged("LHSKeyColumn")
End Property

Public Property Get RHSKeyColumn() As ListColumn
    Set RHSKeyColumn = this.RHSKeyColumn
End Property

Public Property Set RHSKeyColumn(ByVal RHS As ListColumn)
    If Not this.RHSKeyColumn Is Nothing Then
        If this.RHSKeyColumn.Name = RHS.Name Then Exit Property
    End If
    this.IsDirtyQuality = True
    this.IsDirtyMatch = False
    this.IsDirty = True
    Set this.RHSKeyColumn = RHS
    RaiseEvent PropertyChanged("RHSKeyColumn")
End Property

Private Sub UpdateColumns(ByVal lo As ListObject, ByVal coll As Collection)
    ' TODO This fails when selecting new table in key select dialog
    Debug.Assert Not lo Is Nothing
    
    Dim i As Long
    Dim lc As ListColumn
    
    For i = coll.Count To 1 Step -1
        coll.Remove i
    Next i
    
    For Each lc In lo.ListColumns
        coll.Add Item:=lc, key:=lc.Name
    Next lc
End Sub

Private Sub Class_Initialize()
    Set this.LHSColumns = New Collection
    Set this.RHSColumns = New Collection
End Sub

Public Sub TrySelectLHS(ByVal key As String)
    If key <> vbNullString Then
        Set LHSKeyColumn = this.LHSColumns(key)
    End If
End Sub

Public Sub TrySelectRHS(ByVal key As String)
    If key <> vbNullString Then
        Set RHSKeyColumn = this.RHSColumns(key)
    End If
End Sub

Public Sub DoCheck()
    Debug.Assert this.IsDirty = True
    
    'MsgBox "DoWork"
    
    this.IsDirty = False
    
    RaiseEvent CheckCompleted
End Sub

Private Function IViewModel_IsValid() As Boolean
    If this.LHSKeyColumn Is Nothing Then Exit Function
    If this.RHSKeyColumn Is Nothing Then Exit Function
    IViewModel_IsValid = True
End Function

Public Function IsValid() As Boolean
    IsValid = IViewModel_IsValid
End Function

Public Function CanCheck() As Boolean
    If this.LHSTable Is Nothing Then Exit Function
    If this.RHSTable Is Nothing Then Exit Function
    If this.LHSKeyColumn Is Nothing Then Exit Function
    If this.RHSKeyColumn Is Nothing Then Exit Function
    If this.IsDirtyQuality = False Then Exit Function
    CanCheck = True
End Function

Public Function CanMatch() As Boolean
    CanMatch = this.IsDirtyMatch
End Function

Public Function CanContinue() As Boolean
    ' Should only continue if KeyColumnComparer is valid and .Inner.Count > 0
    CanContinue = (this.IsDirtyQuality = False) And (this.IsDirtyMatch = False)
End Function

Public Sub UpdatePreviews()
    this.IsDirtyQuality = False
    this.IsDirtyMatch = True
    RaiseEvent PreviewChanged
End Sub

Public Sub UpdateMatch()
    this.IsDirtyMatch = False
    RaiseEvent MatchChanged
End Sub
