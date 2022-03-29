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
    LHS As ListObject
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
    
    GoBack As Boolean
End Type

Private this As TClsValueMapperViewModel

Public Property Get GoBack() As Boolean
    GoBack = this.GoBack
End Property

Public Property Let GoBack(ByVal RHS As Boolean)
    this.GoBack = RHS
End Property

Public Property Get ShowMappedOnlyLHS() As Boolean
    ShowMappedOnlyLHS = this.ShowMappedOnlyLHS
End Property

Public Property Let ShowMappedOnlyLHS(ByVal RHS As Boolean)
    If this.ShowMappedOnlyLHS = RHS Then Exit Property
    this.ShowMappedOnlyLHS = RHS
    RaiseEvent CollectionChangedLHS
End Property

Public Property Get ShowMappedOnlyRHS() As Boolean
    ShowMappedOnlyRHS = this.ShowMappedOnlyRHS
End Property

Public Property Let ShowMappedOnlyRHS(ByVal RHS As Boolean)
    If this.ShowMappedOnlyRHS = RHS Then Exit Property
    this.ShowMappedOnlyRHS = RHS
    RaiseEvent CollectionChangedRHS
End Property

Public Property Get Flags() As Integer
    Flags = this.Flags
End Property

Public Property Let Flags(ByVal RHS As Integer)
    this.Flags = RHS
End Property

Public Property Get KeyColumnLHS() As ListColumn
    Set KeyColumnLHS = this.KeyColumnLHS
End Property

Public Property Set KeyColumnLHS(ByVal RHS As ListColumn)
    Set this.KeyColumnLHS = RHS
End Property

Public Property Get KeyColumnRHS() As ListColumn
    Set KeyColumnRHS = this.KeyColumnRHS
End Property

Public Property Set KeyColumnRHS(ByVal RHS As ListColumn)
    Set this.KeyColumnRHS = RHS
End Property

Public Property Get checked() As Collection
    Set checked = this.CheckedValuePairs.Item
End Property

Public Property Get LHS() As ListObject
    Set LHS = this.LHS
End Property

Public Property Set LHS(ByVal RHS As ListObject)
    Set this.LHS = RHS
End Property

Public Property Get RHS() As ListObject
    Set RHS = this.RHS
End Property

Public Property Set RHS(ByVal RHS As ListObject)
    Set this.RHS = RHS
End Property

Public Property Get LHSCriteria() As String
    LHSCriteria = this.LHSCriteria
End Property

Public Property Let LHSCriteria(ByVal RHS As String)
    If this.LHSCriteria = RHS Then Exit Property
    this.LHSCriteria = RHS
    RaiseEvent CollectionChangedLHS
End Property

Public Property Get RHSCriteria() As String
    RHSCriteria = this.RHSCriteria
End Property

Public Property Let RHSCriteria(ByVal RHS As String)
    If this.RHSCriteria = RHS Then Exit Property
    this.RHSCriteria = RHS
    RaiseEvent CollectionChangedRHS
End Property

Private Sub Class_Initialize()
    Set this.ValuePairs = New ColumnPairs
    Set this.CheckedValuePairs = New ColumnPairs
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = this.CheckedValuePairs.Count > 0
End Function

Public Sub LoadLHStoListView(ByVal lv As ListView)
    Dim lc As ListColumn
    Dim li As ListItem
    
    lv.ListItems.Clear
    
    For Each lc In LHS.ListColumns
        If (LHSCriteria = vbNullString) Or (lc.Name Like LHSCriteria & "*") Then
            If (this.ShowMappedOnlyLHS = False) Or (Not this.ValuePairs.GetPair(LHS:=lc) Is Nothing) Then
                Set li = lv.ListItems.Add(key:=lc.Name, text:=lc.Name)
                li.ListSubItems.Add text:=GetColumnHeaderFromListColumn(lc)
            End If
        End If
    Next lc
    
    UpdateLHStoListView lv
End Sub

Public Sub UpdateLHStoListView(ByVal lv As ListView)
    Dim li As ListItem
    Dim lc As ListColumn
    
    For Each li In lv.ListItems
        If TryGetListColumnByName(this.LHS, li.key, lc) Then
            If lc Is this.KeyColumnLHS Then
                li.SmallIcon = "Key"
            ElseIf this.ValuePairs.GetPair(LHS:=lc) Is Nothing Then
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
    For Each colPair In this.ValuePairs
        If colPair.LHS.Name = columnName Then
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
            If (this.ShowMappedOnlyRHS = False) Or (this.ValuePairs.IsMappedRHS(lc)) Then
                With lv.ListItems.Add(key:=lc.Name, text:=lc.Name)
                    .ListSubItems.Add text:=vbNullString
                    .ListSubItems.Add text:=GetColumnHeaderFromListColumn(lc)
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
        If TryGetListColumnByName(this.RHS, li.key, lc) Then
            If lc Is this.KeyColumnRHS Then
                li.SmallIcon = "Key"
                With li.ListSubItems(1)
                    .text = KEY_COLUMN
                    .ForeColor = RGB(212, 175, 55)
                End With
            ElseIf this.ValuePairs.IsMappedRHS(lc) Then
                Set mappedTo = this.ValuePairs.GetPair(RHS:=lc)
                li.SmallIcon = "Tick"
                With li.ListSubItems(1)
                    .text = mappedTo.LHS.Name
                    .ForeColor = RGB(0, 0, 0)
                End With
            Else
                li.SmallIcon = "Cross"
                With li.ListSubItems(1)
                    .text = NOT_MAPPED
                    .ForeColor = RGB(128, 128, 128)
                End With
            End If
            
            li.checked = this.CheckedValuePairs.IsMappedRHS(lc)
            
            If ListColumnHasArray(lc) Then
                li.SmallIcon = "Fx"
            End If
        End If
    Next li
End Sub

Private Function GetMappedTo(ByVal columnName As String) As String
    Dim colPair As ColumnPair
    For Each colPair In this.ValuePairs
        If colPair.RHS.Name = columnName Then
            GetMappedTo = colPair.LHS.Name
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
        .ColumnHeaders.Add text:="Column", Width:=100
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .CheckBoxes = False
    End With
    
    If RHS Then
        With lv
            .CheckBoxes = True
            .ColumnHeaders(1).Width = 70
            .ColumnHeaders.Add text:="Mapped to", Width:=128
        End With
    End If
    
    lv.ColumnHeaders.Add text:="R1C1", Width:=40
End Sub

Public Sub TryMap()
    Debug.Assert Not this.SelectedLHS Is Nothing
    Debug.Assert Not this.SelectedRHS Is Nothing
    Dim colPair As ColumnPair
    Set colPair = ColumnPair.Create(this.SelectedLHS, this.SelectedRHS)
    this.ValuePairs.AddOrReplace colPair
    this.CheckedValuePairs.AddOrReplace colPair
    
    this.IsAutoMapped = False
    
    'RaiseEvent CollectionChangedLHS
    RaiseEvent CollectionChangedRHS
    RaiseEvent SelectionChanged
    RaiseEvent MappingChanged
End Sub

Public Sub TryUnMap()
    Dim colPair As ColumnPair
    
    If this.ValuePairs.TryGetPair(colPair, , this.SelectedRHS) Then
        If this.SelectedValuePair Is colPair Then
            Set this.SelectedValuePair = Nothing
        End If
        
        this.CheckedValuePairs.Remove colPair
        this.ValuePairs.Remove colPair
        
        RaiseEvent SelectionChanged
        RaiseEvent MappingChanged
    End If
End Sub

Private Function DoesMappingExist(ByVal colPair As ColumnPair) As Boolean
    Dim RHS As ColumnPair
    For Each RHS In this.ValuePairs
        'If colPair.Equals(rhs) Then
        If colPair.RHS Is RHS.RHS Then
            DoesMappingExist = True
            Exit Function
        End If
    Next RHS
End Function

Private Function IsMappingSelected(ByVal columnName As String) As Boolean
    Dim colPair As ColumnPair
    For Each colPair In this.CheckedValuePairs ' TODO DRY
        If colPair.RHS.Name = columnName Then
            IsMappingSelected = True
            Exit Function
        End If
    Next colPair
End Function

Public Sub TryCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lc As ListColumn
    Dim colPair As ColumnPair
    
    If TryGetListColumnByName(this.RHS, Item.key, lc) Then
        If Item.checked Then
            If this.ValuePairs.IsMappedRHS(lc) Then
                If this.ValuePairs.TryGetPair(colPair:=colPair, RHS:=lc) Then
                    this.CheckedValuePairs.Add colPair
                End If
            Else
                Item.checked = False
            End If
        Else
            If this.ValuePairs.IsMappedRHS(lc) Then
                If this.ValuePairs.TryGetPair(colPair:=colPair, RHS:=lc) Then
                    this.CheckedValuePairs.Remove colPair
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
    If TryGetListColumnByName(this.LHS, Item.key, lc) Then
        Set this.SelectedLHS = lc
    End If

    RaiseEvent SelectionChanged
End Sub

Public Sub TrySelectRHS(ByVal Item As MSComctlLib.ListItem)
    Dim lc As ListColumn
    If TryGetListColumnByName(this.RHS, Item.key, lc) Then
        Set this.SelectedRHS = lc
    End If
    
    Dim colPair As ColumnPair
    Set colPair = this.ValuePairs.GetPair(RHS:=lc)
    Set this.SelectedValuePair = colPair

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
    this.CheckedValuePairs.Clear
    For i = 1 To this.ValuePairs.Count
        this.CheckedValuePairs.Add this.ValuePairs.Item(i)
    Next i
    RaiseEvent MappingChanged
End Sub

Public Sub SelectNone()
    this.CheckedValuePairs.Clear
    RaiseEvent MappingChanged
End Sub

Public Property Get CanMapRight() As Boolean
    If this.SelectedLHS Is this.KeyColumnLHS Then Exit Property
    If this.SelectedRHS Is this.KeyColumnRHS Then Exit Property
    CanMapRight = True
End Property

Public Property Get CanUnmapLeft() As Boolean
    CanUnmapLeft = this.ValuePairs.IsMappedRHS(this.SelectedRHS)
End Property

Public Property Get CanAutoMap() As Boolean
    CanAutoMap = Not this.IsAutoMapped
End Property

Public Property Get CanReset() As Boolean
    CanReset = this.ValuePairs.Count > 0
End Property

Public Property Get CanSelectAll() As Boolean
    CanSelectAll = this.CheckedValuePairs.Count <> this.ValuePairs.Count
End Property

Public Property Get CanSelectNone() As Boolean
    CanSelectNone = this.CheckedValuePairs.Count > 0
End Property

Public Sub Reset()
    this.LHSCriteria = vbNullString
    this.RHSCriteria = vbNullString
    this.ValuePairs.Clear
    this.CheckedValuePairs.Clear
    Set this.SelectedValuePair = Nothing
    this.IsAutoMapped = False
    RaiseEvent SelectionChanged
    'RaiseEvent MappingChanged
    RaiseEvent CollectionChangedLHS
    RaiseEvent CollectionChangedRHS
End Sub

Public Sub Automap()
    Reset
    TryAutomap
End Sub

Public Sub AutomapIfEmpty()
    If this.ValuePairs.Count = 0 Then
        Automap
    End If
End Sub

Private Sub TryAutomap()
    Dim LHS As ListColumn
    Dim RHS As ListColumn
    Dim colPair As ColumnPair
    
    For Each RHS In this.RHS.ListColumns
        If Not RHS Is this.KeyColumnRHS Then
            If TryGetListColumnByName(this.LHS, RHS.Name, LHS) Then
                If Not LHS Is this.KeyColumnLHS Then
                    If ListColumnHasArray(RHS) = False Then
                        Set colPair = ColumnPair.Create(LHS, RHS)
                        this.ValuePairs.Add colPair
                        this.CheckedValuePairs.Add colPair
                    End If
                End If
            End If
        End If
    Next RHS

    this.IsAutoMapped = True
    RaiseEvent MappingChanged
End Sub

Public Sub LoadFromTransferInstruction(ByVal ti As TransferInstruction)
    Dim colPair As ColumnPair
    For Each colPair In ti.ValuePairs
        this.ValuePairs.Add colPair
        this.CheckedValuePairs.Add colPair
    Next colPair
End Sub
