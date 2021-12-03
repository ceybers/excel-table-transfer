VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmValueMapper 
   Caption         =   "Map Value Columns"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8835.001
   OleObjectBlob   =   "frmValueMapper.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmValueMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Completed(arr As Variant)
Public Event Cancelled()

Const NOT_MAPPED As String = "(not mapped)"

Private Type tValueMapper
    src As ListObject
    dst As ListObject
    srcKey As ListColumn
    dstKey As ListColumn
End Type

Private this As tValueMapper

Private Sub cmbAutoMap_Click()
    Dim s As String
    Dim i As Integer, j As Integer
    Dim lv As ListView, lb As MSForms.ListBox
    Set lv = Me.lvListView
    Set lb = Me.lbListBox
    For i = 1 To lv.ListItems.count
        For j = 1 To lb.ListCount
            If CStr(lv.ListItems(i)) = CStr(lb.List(j - 1)) Then
                lv.ListItems(i).Checked = True
                lv.ListItems(i).SubItems(1) = CStr(lb.List(j - 1))
                Me.cmdNext.Enabled = True
                Me.cmdNext.SetFocus
            End If
        Next j
    Next i
End Sub

Private Sub cmbClearSearch_Click()
    Me.txtSearch.text = ""
End Sub

Private Sub cmbSelectAll_Click()
    Dim i As Integer
    For i = 1 To Me.lvListView.ListItems.count
        ' TODO Replace hardcoded value with const
        If Me.lvListView.ListItems(i).SubItems(1) <> "(not mapped)" Then
            Me.lvListView.ListItems(i).Checked = True
        End If
    Next i
End Sub

Private Sub cmbSelectNone_Click()
    Dim i As Integer
    For i = 1 To Me.lvListView.ListItems.count
        Me.lvListView.ListItems(i).Checked = False
    Next i
End Sub

Private Sub cmdNext_Click()
    TryNext
End Sub

Private Sub lbListBox_Click()
    Dim s As String
    s = CStr(Me.lbListBox.List(Me.lbListBox.ListIndex))
    Call TryUpdateInListview(s)
    Call CheckNextButton
End Sub

Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Debug.Print ColumnHeader.text
    
End Sub

Private Sub lvListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call CheckNextButton
End Sub

Private Sub lvListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim s As String
    s = Item.ListSubItems(1)
    
    UnboldAllInListview
    Me.lvListView.ListItems(Item.Index).Bold = True
    Me.lvListView.ListItems(Item.Index).Selected = True
End Sub

Public Sub PopulateColumnMapper(src As ListObject, srcKey As ListColumn, dst As ListObject, dstKey As ListColumn)
    Set this.src = src
    Set this.srcKey = srcKey
    Set this.dst = dst
    Set this.dstKey = dstKey
    
    Call PopulateColumnsToListView(Me.lvListView, src, srcKey)
    Call PopulateColumnsToListBox(Me.lbListBox, dst, dstKey)
End Sub

Private Function PopulateColumnsToListView(lv As ListView, lo As ListObject, keyCol As ListColumn)
    Dim lc As ListColumn
    
    lv.ColumnHeaders.Clear
    lv.ListItems.Clear
    
    With lv
        .View = lvwReport
        .CheckBoxes = True
        .FullRowSelect = True
        .Gridlines = True
    End With
    
    lv.ColumnHeaders.Add , , "Source"
    lv.ColumnHeaders.Add , , "Destination"
    
    Dim li As ListItem
    For Each lc In lo.ListColumns
        If lc <> keyCol Then
            Set li = lv.ListItems.Add()
            With li
                .text = lc.Name
                .ListSubItems.Add , , NOT_MAPPED
            End With
        End If
    Next lc
End Function

Private Function PopulateColumnsToListBox(lb As MSForms.ListBox, lo As ListObject, keyCol As ListColumn, Optional filter As String)
    Dim lc As ListColumn
    Dim currentlySelected As String
    If lb Is Nothing Then Exit Function
    
    If Not (lb = vbNull) Then
        currentlySelected = CStr(lb)
    End If
    
    If lb.ListCount > 0 Then
        lb.Clear
    End If
    
    lb.AddItem NOT_MAPPED
    For Each lc In lo.ListColumns
        If lc <> keyCol Then
            If IsMissing(filter) Or UCase(Left(lc.Name, Len(filter))) = UCase(filter) Then
                lb.AddItem (lc.Name)
            End If
            If currentlySelected <> "" And lc.Name = currentlySelected Then
                lb = lb.List(lb.ListCount - 1)
            End If
        End If
        
    Next lc
End Function

Private Function CheckNextButton()
    Dim li As ListItem
    For Each li In Me.lvListView.ListItems
        If li.Checked = True Then
            Me.cmdNext.Enabled = True
            Exit Function
        End If
    Next li
    Me.cmdNext.Enabled = False
End Function

Private Function TrySelectInListbox(s As String)
    Dim i As Integer
    For i = 0 To Me.lbListBox.ListCount - 1
        If CStr(Me.lbListBox.List(i)) = s Then
            Me.lbListBox.ListIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function UnboldAllInListview()
    Dim li As ListItem
    For Each li In Me.lvListView.ListItems
        li.Bold = False
    Next li
End Function

Private Function TryUpdateInListview(s As String)
    Dim li As ListItem
    Dim selectedLi As ListItem
    For Each li In Me.lvListView.ListItems
        If li.Selected = True Then
            Set selectedLi = li
            selectedLi.ListSubItems(1).text = s
            selectedLi.Checked = (s <> NOT_MAPPED)
            Exit Function
        End If
    Next li
    Exit Function
End Function

Private Sub txtSearch_Change()
    If Len(txtSearch.text) > 0 Then
        PopulateColumnsToListBox Me.lbListBox, this.dst, this.dstKey, txtSearch.text
    Else
        PopulateColumnsToListBox Me.lbListBox, this.dst, this.dstKey
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        RaiseEvent Cancelled
    End If
    'Unload Me
End Sub

Private Function GetArray() As Variant
    Dim i As Integer
    Dim n As Integer
    Dim result As Variant
    
    n = 0
    For i = 1 To Me.lvListView.ListItems.count
        If Me.lvListView.ListItems(i).Checked = True Then
            n = n + 1
        End If
    Next i
    ReDim result(1 To n, 1 To 2)
    
    n = 0
    For i = 1 To Me.lvListView.ListItems.count
        If Me.lvListView.ListItems(i).Checked = True Then
            n = n + 1
            result(n, 1) = Me.lvListView.ListItems(i).text
            result(n, 2) = Me.lvListView.ListItems(i).SubItems(1)
        End If
    Next i
    
    GetArray = result
End Function

Private Sub TryNext()
    Me.Hide
    RaiseEvent Completed(GetArray)
    Unload Me
End Sub
