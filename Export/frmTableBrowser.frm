VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableBrowser 
   Caption         =   "Select a table to transfer data to..."
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "frmTableBrowser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTableBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event TableSelected(ByVal lo As ListObject)
Public Event Cancelled()

Private SelectedURL As String

Private Sub UserForm_Initialize()
    'Call InitialiseForm(Me)
    SetupImages
    PopulateTreeview Me.tvTreeView
End Sub

Private Sub cmbCancel_Click()
    RaiseEvent Cancelled
    Unload Me
End Sub

Private Sub DoSelectTable()
    Me.Hide
    RaiseEvent TableSelected(TableFromString(SelectedURL))
    Unload Me
End Sub

Private Sub cmbOK_Click()
    'Me.Hide
    DoSelectTable
    'Call TransferToTable(TableFromString(SelectedURL))
End Sub

Private Sub tvTreeView_DblClick()
    ' How to get selected node?
    Dim url As String
    url = GetURLFromNode(Me.tvTreeView.SelectedItem)
    If url <> "" Then
        SelectedURL = url
        'Me.Hide
        DoSelectTable
        'Call TransferToTable(TableFromString(SelectedURL))
    End If
End Sub

Private Sub tvTreeView_NodeClick(ByVal node As MSComctlLib.node)
    Dim url As String
    url = GetURLFromNode(node)
    
    If url = "" Then
        Me.cmbOK.Enabled = False
    Else
        SelectedURL = url
        Me.cmbOK.Enabled = True
    End If
End Sub

Private Function GetURLFromNode(ByVal node As MSComctlLib.node) As String
    GetURLFromNode = ""
    Dim parent As MSComctlLib.node
    Dim gparent As MSComctlLib.node
    
    Set parent = node.parent
    If parent Is Nothing Then Exit Function
    
    Set gparent = parent.parent
    If gparent Is Nothing Then Exit Function

    GetURLFromNode = gparent.text & "\" & parent.text & "\" & node.text
End Function

Private Sub SetupImages()
    Dim il As ImageList
    Set il = Me.myImageList
    If il.ListImages.count > 0 Then
        'Debug.Print "Image list count = " & il.ListImages.Count
        Exit Sub
    End If
    
    il.ImageWidth = 16
    il.ImageHeight = 16
    il.ListImages.Add 1, "K001", Me.imgWorkbookRegular.Picture
    il.ListImages.Add 2, "K002", Me.imgWorkbookFill.Picture
    il.ListImages.Add 3, "K003", Me.imgWorksheetRegular.Picture
    il.ListImages.Add 4, "K004", Me.imgWorksheetFill.Picture
    il.ListImages.Add 5, "K005", Me.imgListobjectRegular.Picture
    il.ListImages.Add 6, "K006", Me.imgListobjectFill.Picture
    Set Me.tvTreeView.ImageList = il
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        RaiseEvent Cancelled
    End If
    'Unload Me
End Sub

Private Sub PopulateTreeview(tv As TreeView)
    Dim activeTable As ListObject
    Set activeTable = Selection.ListObject
    
    Dim nds As Nodes
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Integer
    Dim wbIdx As Integer
    Dim wsIdx As Integer
    
    Set nds = tv.Nodes
    
    nds.Clear
    i = 0
    
    For Each wb In Application.Workbooks
        i = i + 1
        nds.Add , , ToKey(i), wb.Name, "K001", "K002"
        nds.Item(nds.count).Expanded = True
        wbIdx = i
        For Each ws In wb.Worksheets
            i = i + 1
            If ws.ListObjects.count > 0 Then
                nds.Add ToKey(wbIdx), tvwChild, ToKey(i), ws.Name, "K003", "K004"
                nds.Item(nds.count).Expanded = True
                wsIdx = i
            End If
            For Each lo In ws.ListObjects
                i = i + 1
                If lo = activeTable And ws.Name = ActiveSheet.Name And wb.Name = ActiveWorkbook.Name Then
                    nds.Add ToKey(wsIdx), tvwChild, ToKey(i), lo.Name & " (active)", "K006", "K006"
                Else
                    nds.Add ToKey(wsIdx), tvwChild, ToKey(i), lo.Name, "K005", "K005"
                End If
            Next lo
        Next ws
    Next wb
End Sub


