VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKeyMapper 
   Caption         =   "Map Key Columns"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   OleObjectBlob   =   "frmKeyMapper.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKeyMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Completed(lhs As ListColumn, RHS As ListColumn)
Public Event Cancelled()

Private Type tKeyMapper
    Source As ListObject
    Destination As ListObject
    SourceColumn As ListColumn
    DestinationColumn As ListColumn
End Type

Private this As tKeyMapper

Public Function SetTables(Source As ListObject, Destination As ListObject)
    Set this.Source = Source
    Set this.Destination = Destination
    
    'LoadColumnsToComboBox this.Source, Me.cmbSource
    LoadColumnsToImageCombo2 this.Source, Me.cmbSource2
    'LoadColumnsToComboBox this.Destination, Me.cmbDestination
    LoadColumnsToImageCombo2 this.Destination, Me.cmbDestination2
    
    CanCheckNow
    DoCheckNow Fast:=True
End Function

Private Function LoadColumnsToComboBox(lo As ListObject, cmb As MSForms.ComboBox)
    cmb.Clear
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        cmb.AddItem lc.Name, lc.Index - 1
    Next lc
    
    If cmb.ListCount > 0 Then
        cmb = cmb.List(0)
    End If
End Function

Private Function LoadColumnsToImageCombo2(lo As ListObject, cmb As ImageCombo2)
    cmb.ComboItems.Clear
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        cmb.ComboItems.Add Index:=lc.Index, text:=lc.Name
    Next lc
    
    If cmb.ComboItems.count > 0 Then
        cmb.ComboItems(1).Selected = True
    End If
End Function

Private Sub CanCheckNow()
    'If Me.cmbSource.ListIndex = -1 Or _
       'Me.cmbDestination.ListIndex = -1 Then
        'Me.cmbCheck.Enabled = False
    'Else
        'Me.cmbCheck.Enabled = True
    'End If
    
    'Me.cmbMatchToDst.Enabled = CanMatchDst
    'Me.cmbMatchToSrc.Enabled = CanMatchSrc
    Me.cmbMatchToDst.Enabled = CanMatchImageCombo2(Me.cmbSource2, Me.cmbDestination2)
    Me.cmbMatchToSrc.Enabled = CanMatchImageCombo2(Me.cmbDestination2, Me.cmbSource2)
    Me.cmbCheck.Enabled = True
End Sub

Private Sub cmbBack_Click()
    RaiseEvent Cancelled
    Unload Me
End Sub

Private Sub cmbCheck_Click()
    DoCheckNow
    Me.cmbCheck.Enabled = False
End Sub

Private Sub cmbMatchToDst_Click()
    'DoMatchDst
    DoMatchImageCombo2 Me.cmbSource2, Me.cmbDestination2
    Me.cmbCheck.SetFocus
    Me.cmbMatchToDst.Enabled = False
    Me.cmbMatchToSrc.Enabled = False
End Sub

Private Sub cmbMatchToSrc_Click()
    'DoMatchSrc
    DoMatchImageCombo2 Me.cmbDestination2, Me.cmbSource2
    Me.cmbCheck.SetFocus
    Me.cmbMatchToDst.Enabled = False
    Me.cmbMatchToSrc.Enabled = False
End Sub

Private Function CanMatchImageCombo2(lhs As ImageCombo2, RHS As ImageCombo) As Boolean
    CanMatchImageCombo2 = False
    If lhs.ComboItems.count = 0 Or RHS.ComboItems.count = 0 Then
        Exit Function
    End If
    Dim lookFor As String, curSel As String, thisItem As String
    lookFor = lhs.SelectedItem.text
    curSel = RHS.SelectedItem.text
    If lookFor = curSel Then Exit Function
    Dim i As Integer
    For i = 1 To RHS.ComboItems.count
        thisItem = RHS.ComboItems(i).text
        If lookFor = thisItem Then
            If (curSel <> thisItem) Then
                CanMatchImageCombo2 = True
            End If
            Exit Function
        End If
    Next i
End Function

Private Function DoMatchImageCombo2(lhs As ImageCombo2, RHS As ImageCombo)
    Dim lookFor As String, thisItem As String
    lookFor = lhs.SelectedItem.text
    Dim i As Integer
    For i = 1 To RHS.ComboItems.count
        thisItem = RHS.ComboItems(i).text
        If lookFor = thisItem Then
            RHS.ComboItems(i).Selected = True
            Exit Function
        End If
    Next i
End Function

Private Sub cmbNext_Click()
    Me.Hide
    RaiseEvent Completed(this.SourceColumn, this.DestinationColumn)
    Unload Me
End Sub

Private Sub cmbDestination2_Click()
    CanCheckNow
End Sub

Private Sub cmbDestination2_Dropdown()
    CanCheckNow
End Sub

Private Sub cmbDestination2_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
    CanCheckNow
End Sub

Private Sub cmbSource2_Click()
    CanCheckNow
End Sub

Private Sub cmbSource2_Dropdown()
    CanCheckNow
End Sub

Private Sub cmbSource2_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
    CanCheckNow
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        RaiseEvent Cancelled
    End If
End Sub

Private Sub DoCheckNow(Optional Fast As Boolean = False)
    Dim lhs As Variant
    Dim RHS As Variant
    
    'Set this.SourceColumn = this.Source.ListColumns(CStr(Me.cmbSource))
    'Set this.DestinationColumn = this.Destination.ListColumns(CStr(Me.cmbDestination))
    
    Set this.SourceColumn = this.Source.ListColumns(Me.cmbSource2.SelectedItem.text)
    Set this.DestinationColumn = this.Destination.ListColumns(Me.cmbDestination2.SelectedItem.text)
    
    lhs = this.SourceColumn.DataBodyRange.Value
    RHS = this.DestinationColumn.DataBodyRange.Value
    
    If (UBound(lhs, 1) > 1000) Or (UBound(RHS, 1) > 1000) Then
        If Fast = True Then
            Exit Sub
        Else
            If vbYes = MsgBox("More than 1000 items! Assume OK?", vbYesNo) Then
                TryNextNow Force:=True
                Exit Sub
            End If
        End If
    End If
    'lhs = ArrayTrim(lhs, 100)
    'rhs = ArrayTrim(rhs, 100)
    
    With ArrayAnalyseTwo(lhs, RHS)
        LoadArrayToListBox ArrayTrim(.LeftOnly, 100), Me.lbLHSOnly
        LoadArrayToListBox ArrayTrim(.Intersection, 100), Me.lbIntersect
        LoadArrayToListBox ArrayTrim(.RightOnly, 100), Me.lbRHSOnly
        
        LoadArrayCountToLabel .LeftOnlyCount, Me.lblLHSMatches, " additions(s)"
        LoadArrayCountToLabel .IntersectionCount, Me.lblIntMatches, " matches(s)"
        LoadArrayCountToLabel .RightOnlyCount, Me.lblRHSMatches, " removals(s)"
    End With

    With ArrayAnalyseOne(lhs)
        LoadArrayCountToLabel .Errors, Me.lblLHSErrors, " error cell(s)"
        LoadArrayCountToLabel .Blanks, Me.lblLHSBlanks, " blank cell(s)"
    End With

    With ArrayAnalyseOne(RHS)
        LoadArrayCountToLabel .Errors, Me.lblRHSErrors, " error cell(s)"
        LoadArrayCountToLabel .Blanks, Me.lblRHSBlanks, " blank cell(s)"
    End With
    
    TryNextNow
End Sub

Private Sub TryNextNow(Optional Force As Boolean = False)
    If (Force = True) Or (Me.lbIntersect.ListCount > 0) Then
        Me.cmbNext.Enabled = True
        Me.cmbNext.SetFocus
    Else
        Me.cmbNext.Enabled = False
    End If
End Sub

Private Function LoadArrayCountToLabel(n As Integer, lbl As MSForms.Label, suffix As String)
    lbl.Caption = CStr(n) & suffix
End Function

Private Function LoadArrayToListBox(arr As Variant, lb As MSForms.ListBox)
    Dim i As Integer
    lb.Clear
    For i = 1 To UBound(arr, 1)
        lb.AddItem arr(i, 1)
    Next i
End Function


