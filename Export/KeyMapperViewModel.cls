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

Private This As TViewModel

Public Property Get IsDirty() As Boolean
    IsDirty = This.IsDirty
End Property

Public Property Get LHSTable() As ListObject
    Set LHSTable = This.LHSTable
End Property

Public Property Set LHSTable(ByVal RHS As ListObject)
    If This.LHSTable Is RHS Then Exit Property
    Set This.LHSTable = RHS
    This.IsDirty = True
    Debug.Assert Not RHS Is Nothing
    UpdateColumns This.LHSTable, This.LHSColumns
    RaiseEvent PropertyChanged("LHSTable")
    RaiseEvent PropertyChanged("LHSColumns")
End Property

Public Property Get RHSTable() As ListObject
    Set RHSTable = This.RHSTable
End Property

Public Property Set RHSTable(ByVal RHS As ListObject)
    If This.RHSTable Is RHS Then Exit Property
    Set This.RHSTable = RHS
    This.IsDirty = True
    UpdateColumns This.RHSTable, This.RHSColumns
    RaiseEvent PropertyChanged("RHSTable")
    RaiseEvent PropertyChanged("RHSColumns")
End Property

Public Property Get LHSColumns() As Collection
    Set LHSColumns = This.LHSColumns
End Property

Public Property Set LHSColumns(ByVal RHS As Collection)
    Set This.LHSColumns = RHS
End Property

Public Property Get RHSColumns() As Collection
    Set RHSColumns = This.RHSColumns
End Property

Public Property Set RHSColumns(ByVal RHS As Collection)
    Set This.RHSColumns = RHS
End Property

Public Property Get LHSKeyColumn() As ListColumn
    Set LHSKeyColumn = This.LHSKeyColumn
End Property

Public Property Set LHSKeyColumn(ByVal RHS As ListColumn)
    If Not This.LHSKeyColumn Is Nothing Then
        If This.LHSKeyColumn.Name = RHS.Name Then Exit Property
    End If
    This.IsDirtyQuality = True
    This.IsDirtyMatch = False
    This.IsDirty = True
    Set This.LHSKeyColumn = RHS
    RaiseEvent PropertyChanged("LHSKeyColumn")
End Property

Public Property Get RHSKeyColumn() As ListColumn
    Set RHSKeyColumn = This.RHSKeyColumn
End Property

Public Property Set RHSKeyColumn(ByVal RHS As ListColumn)
    If Not This.RHSKeyColumn Is Nothing Then
        If This.RHSKeyColumn.Name = RHS.Name Then Exit Property
    End If
    This.IsDirtyQuality = True
    This.IsDirtyMatch = False
    This.IsDirty = True
    Set This.RHSKeyColumn = RHS
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
    Set This.LHSColumns = New Collection
    Set This.RHSColumns = New Collection
End Sub

Public Sub TrySelectLHS(ByVal key As String)
    If key <> vbNullString Then
        Set LHSKeyColumn = This.LHSColumns(key)
    End If
End Sub

Public Sub TrySelectRHS(ByVal key As String)
    If key <> vbNullString Then
        Set RHSKeyColumn = This.RHSColumns(key)
    End If
End Sub

Public Sub DoCheck()
    Debug.Assert This.IsDirty = True
    
    'MsgBox "DoWork"
    
    This.IsDirty = False
    
    RaiseEvent CheckCompleted
End Sub

Private Function IViewModel_IsValid() As Boolean
    If This.LHSKeyColumn Is Nothing Then Exit Function
    If This.RHSKeyColumn Is Nothing Then Exit Function
    IViewModel_IsValid = True
End Function

Public Function IsValid() As Boolean
    IsValid = IViewModel_IsValid
End Function

Public Function CanCheck() As Boolean
    If This.LHSTable Is Nothing Then Exit Function
    If This.RHSTable Is Nothing Then Exit Function
    If This.LHSKeyColumn Is Nothing Then Exit Function
    If This.RHSKeyColumn Is Nothing Then Exit Function
    If This.IsDirtyQuality = False Then Exit Function
    CanCheck = True
End Function

Public Function CanMatch() As Boolean
    CanMatch = This.IsDirtyMatch
End Function

Public Function CanContinue() As Boolean
    ' Should only continue if KeyColumnComparer is valid and .Inner.Count > 0
    CanContinue = (This.IsDirtyQuality = False) And (This.IsDirtyMatch = False)
End Function

Public Sub UpdatePreviews()
    This.IsDirtyQuality = False
    This.IsDirtyMatch = True
    RaiseEvent PreviewChanged
End Sub

Public Sub UpdateMatch()
    This.IsDirtyMatch = False
    RaiseEvent MatchChanged
End Sub
