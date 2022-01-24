VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsValueMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ValueMapper2")
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
    ValuePairs As clsColumnPairs
    CheckedValuePairs As clsColumnPairs
    SelectedLHS As ListColumn
    SelectedRHS As ListColumn
    SelectedValuePair As clsColumnPair
    IsReset As Boolean
    IsAutoMapped As Boolean
End Type

Private this As TClsValueMapperViewModel

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

Public Property Get Checked() As Collection
    Set Checked = this.CheckedValuePairs.Item
End Property

Public Property Get lhs() As ListObject
    Set lhs = this.lhs
End Property

Public Property Set lhs(ByVal RHS As ListObject)
    Set this.lhs = RHS
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
    Set this.ValuePairs = New clsColumnPairs
    Set this.CheckedValuePairs = New clsColumnPairs
    this.IsReset = True
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = this.CheckedValuePairs.Count > 0
End Function

Public Sub LoadLHStoListView(ByRef lv As ListView)
    Dim lc As ListColumn
    Dim li As ListItem
    
    lv.ListItems.Clear
    
    For Each lc In lhs.ListColumns
        If (LHSCriteria = vbNullString) Or (lc.Name Like LHSCriteria & "*") Then
            Set li = lv.ListItems.Add(key:=lc.Name, text:=lc.Name)
        End If
    Next lc
    
    UpdateLHStoListView lv
End Sub

Public Sub UpdateLHStoListView(ByRef lv As ListView)
    Dim li As ListItem
    Dim lc As ListColumn
    
    For Each li In lv.ListItems
        If TryGetListColumnByName(this.lhs, li.key, lc) Then
            If lc Is this.KeyColumnLHS Then
                li.SmallIcon = "Key"
            ElseIf this.ValuePairs.GetPair(lhs:=lc) Is Nothing Then
                li.SmallIcon = "Cross"
            Else
                li.SmallIcon = "Tick"
            End If
        End If
    Next li
End Sub

Private Function IsMappedLHS(ByRef columnName As String) As Boolean
    Err.Raise 5, , "no"
    Dim colPair As clsColumnPair
    For Each colPair In this.ValuePairs
        If colPair.lhs.Name = columnName Then
            IsMappedLHS = True
            Exit Function
        End If
    Next colPair
End Function

Public Sub LoadRHStoListView(ByRef lv As ListView)
    Dim lc As ListColumn
    
    lv.ListItems.Clear
    
    For Each lc In RHS.ListColumns
        If RHSCriteria = vbNullString Or lc.Name Like RHSCriteria & "*" Then
            With lv.ListItems.Add(key:=lc.Name, text:=lc.Name)
                .ListSubItems.Add text:=vbNullString
            End With
        End If
    Next lc
End Sub

Public Sub UpdateRHStoListView(ByRef lv As ListView)
    Dim lc As ListColumn
    Dim li As ListItem
    Dim lsi As ListSubItem
    Dim subItem As String
    Dim mappedTo As clsColumnPair
    
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
            
            li.Checked = this.CheckedValuePairs.IsMappedRHS(lc)
        End If
    Next li
End Sub

Private Function GetMappedTo(ByRef columnName As String) As String
    Dim colPair As clsColumnPair
    For Each colPair In this.ValuePairs
        If colPair.RHS.Name = columnName Then
            GetMappedTo = colPair.lhs.Name
            Exit Function
        End If
    Next colPair
End Function

Private Function IsColumnMapped(ByRef columnName As String) As Boolean
    IsColumnMapped = Not GetMappedTo(columnName) = vbNullString
End Function

Public Sub InitializeListView(ByRef lv As ListView, Optional RHS As Boolean = False)
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
    Debug.Assert Not this.SelectedLHS Is Nothing
    Debug.Assert Not this.SelectedRHS Is Nothing
    Dim colPair As clsColumnPair
    Set colPair = clsColumnPair.Create(this.SelectedLHS, this.SelectedRHS)
    this.ValuePairs.AddOrReplace colPair
    this.CheckedValuePairs.AddOrReplace colPair
    
    this.IsReset = False
    this.IsAutoMapped = False
    
    'RaiseEvent CollectionChangedLHS
    RaiseEvent CollectionChangedRHS
    RaiseEvent SelectionChanged
    RaiseEvent MappingChanged
End Sub

Public Sub TryUnMap()
    Dim colPair As clsColumnPair
    
    If this.ValuePairs.TryGetPair(colPair, , this.SelectedRHS) Then
        If this.SelectedValuePair Is colPair Then
            this.SelectedValuePair = Nothing
        End If
        
        this.CheckedValuePairs.Remove colPair
        this.ValuePairs.Remove colPair
        
        RaiseEvent SelectionChanged
        RaiseEvent MappingChanged
    End If
End Sub

Private Function DoesMappingExist(ByRef colPair As clsColumnPair) As Boolean
    Dim RHS As clsColumnPair
    For Each RHS In this.ValuePairs
        'If colPair.Equals(rhs) Then
        If colPair.RHS Is RHS.RHS Then
            DoesMappingExist = True
            Exit Function
        End If
    Next RHS
End Function

Private Function IsMappingSelected(ByRef columnName As String) As Boolean
    Dim colPair As clsColumnPair
    For Each colPair In this.CheckedValuePairs ' TODO DRY
        If colPair.RHS.Name = columnName Then
            IsMappingSelected = True
            Exit Function
        End If
    Next colPair
End Function

Public Sub TryCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lc As ListColumn
    Dim colPair As clsColumnPair
    
    If TryGetListColumnByName(this.RHS, Item.key, lc) Then
        If Item.Checked Then
            If this.ValuePairs.IsMappedRHS(lc) Then
                If this.ValuePairs.TryGetPair(colPair:=colPair, RHS:=lc) Then
                    this.CheckedValuePairs.Add colPair
                End If
            Else
                Item.Checked = False
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
    If TryGetListColumnByName(this.lhs, Item.key, lc) Then
        Set this.SelectedLHS = lc
    End If

    RaiseEvent SelectionChanged
End Sub

Public Sub TrySelectRHS(ByVal Item As MSComctlLib.ListItem)
    Dim lc As ListColumn
    If TryGetListColumnByName(this.RHS, Item.key, lc) Then
        Set this.SelectedRHS = lc
    End If
    
    Dim colPair As clsColumnPair
    Set colPair = this.ValuePairs.GetPair(colPair, RHS:=lc)
    Set this.SelectedValuePair = colPair

    RaiseEvent SelectionChanged
End Sub

' TODO Helper func
Private Function TryGetListColumnByName(ByRef lo As ListObject, ByVal columnName As String, ByRef lc As ListColumn) As Boolean
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

Public Property Get CanMapRight()
    If this.SelectedLHS Is this.KeyColumnLHS Then Exit Property
    If this.SelectedRHS Is this.KeyColumnRHS Then Exit Property
    CanMapRight = True
End Property

Public Property Get CanUnmapLeft()
    CanUnmapLeft = this.ValuePairs.IsMappedRHS(this.SelectedRHS)
End Property

Public Property Get CanAutoMap()
    CanAutoMap = Not this.IsAutoMapped
End Property

Public Property Get CanReset()
    'CanReset = Not this.IsReset ' TODO Not necessary, just check count
    CanReset = this.ValuePairs.Count > 0
End Property

Public Property Get CanSelectAll()
    CanSelectAll = this.CheckedValuePairs.Count <> this.ValuePairs.Count
End Property

Public Property Get CanSelectNone()
    CanSelectNone = this.CheckedValuePairs.Count > 0
End Property

Public Sub Reset()
    this.LHSCriteria = vbNullString
    this.RHSCriteria = vbNullString
    this.ValuePairs.Clear
    this.CheckedValuePairs.Clear
    Set this.SelectedValuePair = Nothing
    this.IsReset = True
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

Private Sub TryAutomap()
    Dim lhs As ListColumn
    Dim RHS As ListColumn
    Dim colPair As clsColumnPair
    
    For Each RHS In this.RHS.ListColumns
        If TryGetListColumnByName(this.lhs, RHS.Name, lhs) Then
            Set colPair = clsColumnPair.Create(lhs, RHS)
            this.ValuePairs.Add colPair
            this.CheckedValuePairs.Add colPair
        End If
    Next RHS
    
    this.IsAutoMapped = True
    this.IsReset = False
    RaiseEvent MappingChanged
End Sub
