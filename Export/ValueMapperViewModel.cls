VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ValueMapper"
Option Explicit
Option Compare Text
Implements IViewModel

Public Event CollectionChangedLHS()
Public Event CollectionChangedRHS()
Public Event MappingChanged()
Public Event SelectionChanged()

Private Const NOT_MAPPED As String = "(Not mapped)"
Private Const KEY_COLUMN As String = "(Key column)"

Private Type TClsValueMapperViewModel
    lhs As ListObject
    RHS As ListObject
    LHSCriteria As String
    RHSCriteria As String
    KeyColumnLHS As ListColumn
    KeyColumnRHS As ListColumn
    'ValuePairs As Collection
    'SelectedValuePairs As Collection
    ValuePairs As ColumnPairs
    CheckedValuePairs As ColumnPairs
    SelectedLHS As ListColumn
    SelectedRHS As ListColumn
    SelectedValuePair As ColumnPair
    IsAutoMapped As Boolean
    Flags As Integer
    
    ShowMappedOnlyLHS As Boolean
    ShowMappedOnlyRHS As Boolean
End Type

Private This As TClsValueMapperViewModel

Public Property Get ShowMappedOnlyLHS() As Boolean
    ShowMappedOnlyLHS = This.ShowMappedOnlyLHS
End Property

Public Property Let ShowMappedOnlyLHS(ByVal RHS As Boolean)
    If This.ShowMappedOnlyLHS = RHS Then Exit Property
    This.ShowMappedOnlyLHS = RHS
    RaiseEvent CollectionChangedLHS
End Property

Public Property Get ShowMappedOnlyRHS() As Boolean
    ShowMappedOnlyRHS = This.ShowMappedOnlyRHS
End Property

Public Property Let ShowMappedOnlyRHS(ByVal RHS As Boolean)
    If This.ShowMappedOnlyRHS = RHS Then Exit Property
    This.ShowMappedOnlyRHS = RHS
    RaiseEvent CollectionChangedRHS
End Property

Public Property Get Flags() As Integer
    Flags = This.Flags
End Property

Public Property Let Flags(ByVal RHS As Integer)
    This.Flags = RHS
End Property

Public Property Get KeyColumnLHS() As ListColumn
    Set KeyColumnLHS = This.KeyColumnLHS
End Property

Public Property Set KeyColumnLHS(ByVal RHS As ListColumn)
    Set This.KeyColumnLHS = RHS
End Property

Public Property Get KeyColumnRHS() As ListColumn
    Set KeyColumnRHS = This.KeyColumnRHS
End Property

Public Property Set KeyColumnRHS(ByVal RHS As ListColumn)
    Set This.KeyColumnRHS = RHS
End Property

Public Property Get checked() As Collection
    Set checked = This.CheckedValuePairs.Item
End Property

Public Property Get lhs() As ListObject
    Set lhs = This.lhs
End Property

Public Property Set lhs(ByVal RHS As ListObject)
    Set This.lhs = RHS
End Property

Public Property Get RHS() As ListObject
    Set RHS = This.RHS
End Property

Public Property Set RHS(ByVal RHS As ListObject)
    Set This.RHS = RHS
End Property

Public Property Get LHSCriteria() As String
    LHSCriteria = This.LHSCriteria
End Property

Public Property Let LHSCriteria(ByVal RHS As String)
    If This.LHSCriteria = RHS Then Exit Property
    This.LHSCriteria = RHS
    RaiseEvent CollectionChangedLHS
End Property

Public Property Get RHSCriteria() As String
    RHSCriteria = This.RHSCriteria
End Property

Public Property Let RHSCriteria(ByVal RHS As String)
    If This.RHSCriteria = RHS Then Exit Property
    This.RHSCriteria = RHS
    RaiseEvent CollectionChangedRHS
End Property

Private Sub Class_Initialize()
    Set This.ValuePairs = New ColumnPairs
    Set This.CheckedValuePairs = New ColumnPairs
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = This.CheckedValuePairs.Count > 0
End Function

Public Sub LoadLHStoListView(ByVal lv As ListView)
    Dim lc As ListColumn
    Dim li As ListItem
    
    lv.ListItems.Clear
    
    For Each lc In lhs.ListColumns
        If (LHSCriteria = vbNullString) Or (lc.Name Like LHSCriteria & "*") Then
            If (This.ShowMappedOnlyLHS = False) Or (Not This.ValuePairs.GetPair(lhs:=lc) Is Nothing) Then
                Set li = lv.ListItems.Add(key:=lc.Name, text:=lc.Name)
            End If
        End If
    Next lc
    
    UpdateLHStoListView lv
End Sub

Public Sub UpdateLHStoListView(ByVal lv As ListView)
    Dim li As ListItem
    Dim lc As ListColumn
    
    For Each li In lv.ListItems
        If TryGetListColumnByName(This.lhs, li.key, lc) Then
            If lc Is This.KeyColumnLHS Then
                li.SmallIcon = "Key"
            ElseIf This.ValuePairs.GetPair(lhs:=lc) Is Nothing Then
                li.SmallIcon = "Cross"
            Else
                li.SmallIcon = "Tick"
            End If
        End If
    Next li
End Sub

Private Function IsMappedLHS(ByVal columnName As String) As Boolean
    Err.Raise 5, , "no"
    Dim colPair As ColumnPair
    For Each colPair In This.ValuePairs
        If colPair.lhs.Name = columnName Then
            IsMappedLHS = True
            Exit Function
        End If
    Next colPair
End Function

Public Sub LoadRHStoListView(ByVal lv As ListView)
    Dim lc As ListColumn
    
    lv.ListItems.Clear
    
    For Each lc In RHS.ListColumns
        If RHSCriteria = vbNullString Or lc.Name Like RHSCriteria & "*" Then
            If (This.ShowMappedOnlyRHS = False) Or (This.ValuePairs.IsMappedRHS(lc)) Then
                With lv.ListItems.Add(key:=lc.Name, text:=lc.Name)
                    .ListSubItems.Add text:=vbNullString
                End With
            End If
        End If
    Next lc
End Sub

Public Sub UpdateRHStoListView(ByVal lv As ListView)
    Dim lc As ListColumn
    Dim li As ListItem
    Dim lsi As ListSubItem
    Dim subItem As String
    Dim mappedTo As ColumnPair
    
    For Each li In lv.ListItems
        If TryGetListColumnByName(This.RHS, li.key, lc) Then
            If lc Is This.KeyColumnRHS Then
                li.SmallIcon = "Key"
                With li.ListSubItems(1)
                    .text = KEY_COLUMN
                    .ForeColor = RGB(212, 175, 55)
                End With
            ElseIf This.ValuePairs.IsMappedRHS(lc) Then
                Set mappedTo = This.ValuePairs.GetPair(RHS:=lc)
                li.SmallIcon = "Tick"
                With li.ListSubItems(1)
                    .text = mappedTo.lhs.Name
                    .ForeColor = RGB(0, 0, 0)
                End With
            Else
                li.SmallIcon = "Cross"
                With li.ListSubItems(1)
                    .text = NOT_MAPPED
                    .ForeColor = RGB(128, 128, 128)
                End With
            End If
            
            li.checked = This.CheckedValuePairs.IsMappedRHS(lc)
        End If
    Next li
End Sub

Private Function GetMappedTo(ByVal columnName As String) As String
    Dim colPair As ColumnPair
    For Each colPair In This.ValuePairs
        If colPair.RHS.Name = columnName Then
            GetMappedTo = colPair.lhs.Name
            Exit Function
        End If
    Next colPair
End Function

Private Function IsColumnMapped(ByVal columnName As String) As Boolean
    IsColumnMapped = Not GetMappedTo(columnName) = vbNullString
End Function

Public Sub InitializeListView(ByVal lv As MSComctlLib.ListView, Optional ByVal RHS As Boolean = False)
    With lv
        .Appearance = cc3D
        .BorderStyle = ccNone
        .view = lvwReport
        .Gridlines = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:="Column", Width:=140
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .CheckBoxes = False
    End With
    If RHS Then
        With lv
            .CheckBoxes = True
            .ColumnHeaders(1).Width = 70
            .ColumnHeaders.Add text:="Mapped to", Width:=68
        End With
    End If
End Sub

Public Sub TryMap()
    Debug.Assert Not This.SelectedLHS Is Nothing
    Debug.Assert Not This.SelectedRHS Is Nothing
    Dim colPair As ColumnPair
    Set colPair = ColumnPair.Create(This.SelectedLHS, This.SelectedRHS)
    This.ValuePairs.AddOrReplace colPair
    This.CheckedValuePairs.AddOrReplace colPair
    
    This.IsAutoMapped = False
    
    'RaiseEvent CollectionChangedLHS
    RaiseEvent CollectionChangedRHS
    RaiseEvent SelectionChanged
    RaiseEvent MappingChanged
End Sub

Public Sub TryUnMap()
    Dim colPair As ColumnPair
    
    If This.ValuePairs.TryGetPair(colPair, , This.SelectedRHS) Then
        If This.SelectedValuePair Is colPair Then
            Set This.SelectedValuePair = Nothing
        End If
        
        This.CheckedValuePairs.Remove colPair
        This.ValuePairs.Remove colPair
        
        RaiseEvent SelectionChanged
        RaiseEvent MappingChanged
    End If
End Sub

Private Function DoesMappingExist(ByVal colPair As ColumnPair) As Boolean
    Dim RHS As ColumnPair
    For Each RHS In This.ValuePairs
        'If colPair.Equals(rhs) Then
        If colPair.RHS Is RHS.RHS Then
            DoesMappingExist = True
            Exit Function
        End If
    Next RHS
End Function

Private Function IsMappingSelected(ByVal columnName As String) As Boolean
    Dim colPair As ColumnPair
    For Each colPair In This.CheckedValuePairs ' TODO DRY
        If colPair.RHS.Name = columnName Then
            IsMappingSelected = True
            Exit Function
        End If
    Next colPair
End Function

Public Sub TryCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lc As ListColumn
    Dim colPair As ColumnPair
    
    If TryGetListColumnByName(This.RHS, Item.key, lc) Then
        If Item.checked Then
            If This.ValuePairs.IsMappedRHS(lc) Then
                If This.ValuePairs.TryGetPair(colPair:=colPair, RHS:=lc) Then
                    This.CheckedValuePairs.Add colPair
                End If
            Else
                Item.checked = False
            End If
        Else
            If This.ValuePairs.IsMappedRHS(lc) Then
                If This.ValuePairs.TryGetPair(colPair:=colPair, RHS:=lc) Then
                    This.CheckedValuePairs.Remove colPair
                Else
                    Debug.Print "TryCheck Failed"
                End If
            End If
        End If
    End If
    
    RaiseEvent MappingChanged
End Sub

Public Sub TrySelectLHS(ByVal Item As MSComctlLib.ListItem)
    Dim lc As ListColumn
    If TryGetListColumnByName(This.lhs, Item.key, lc) Then
        Set This.SelectedLHS = lc
    End If

    RaiseEvent SelectionChanged
End Sub

Public Sub TrySelectRHS(ByVal Item As MSComctlLib.ListItem)
    Dim lc As ListColumn
    If TryGetListColumnByName(This.RHS, Item.key, lc) Then
        Set This.SelectedRHS = lc
    End If
    
    Dim colPair As ColumnPair
    Set colPair = This.ValuePairs.GetPair(RHS:=lc)
    Set This.SelectedValuePair = colPair

    RaiseEvent SelectionChanged
End Sub

' TODO Helper func
Private Function TryGetListColumnByName(ByVal lo As ListObject, ByVal columnName As String, ByRef lc As ListColumn) As Boolean
    Dim curCol As ListColumn
    For Each curCol In lo.ListColumns
        If curCol.Name = columnName Then
            Set lc = curCol
            TryGetListColumnByName = True
            Exit Function
        End If
    Next curCol
End Function

Public Sub SelectAll()
    ' TODO Helper
    Dim i As Long
    This.CheckedValuePairs.Clear
    For i = 1 To This.ValuePairs.Count
        This.CheckedValuePairs.Add This.ValuePairs.Item(i)
    Next i
    RaiseEvent MappingChanged
End Sub

Public Sub SelectNone()
    This.CheckedValuePairs.Clear
    RaiseEvent MappingChanged
End Sub

Public Property Get CanMapRight() As Boolean
    If This.SelectedLHS Is This.KeyColumnLHS Then Exit Property
    If This.SelectedRHS Is This.KeyColumnRHS Then Exit Property
    CanMapRight = True
End Property

Public Property Get CanUnmapLeft() As Boolean
    CanUnmapLeft = This.ValuePairs.IsMappedRHS(This.SelectedRHS)
End Property

Public Property Get CanAutoMap() As Boolean
    CanAutoMap = Not This.IsAutoMapped
End Property

Public Property Get CanReset() As Boolean
    CanReset = This.ValuePairs.Count > 0
End Property

Public Property Get CanSelectAll() As Boolean
    CanSelectAll = This.CheckedValuePairs.Count <> This.ValuePairs.Count
End Property

Public Property Get CanSelectNone() As Boolean
    CanSelectNone = This.CheckedValuePairs.Count > 0
End Property

Public Sub Reset()
    This.LHSCriteria = vbNullString
    This.RHSCriteria = vbNullString
    This.ValuePairs.Clear
    This.CheckedValuePairs.Clear
    Set This.SelectedValuePair = Nothing
    This.IsAutoMapped = False
    RaiseEvent SelectionChanged
    'RaiseEvent MappingChanged
    RaiseEvent CollectionChangedLHS
    RaiseEvent CollectionChangedRHS
End Sub

Public Sub Automap()
    Reset
    TryAutomap
End Sub

Private Sub TryAutomap()
    Dim lhs As ListColumn
    Dim RHS As ListColumn
    Dim colPair As ColumnPair
    
    For Each RHS In This.RHS.ListColumns
        If Not RHS Is This.KeyColumnRHS Then
            If TryGetListColumnByName(This.lhs, RHS.Name, lhs) Then
                If Not lhs Is This.KeyColumnLHS Then
                    Set colPair = ColumnPair.Create(lhs, RHS)
                    This.ValuePairs.Add colPair
                    This.CheckedValuePairs.Add colPair
                End If
            End If
        End If
    Next RHS

    This.IsAutoMapped = True
    RaiseEvent MappingChanged
End Sub
