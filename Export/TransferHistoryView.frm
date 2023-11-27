VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferHistoryView 
   Caption         =   "Table Transfer History"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "TransferHistoryView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransferHistoryView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'@Folder "TransferHistory"
Option Explicit
Implements IView

Private Type TView
    ' Context as MVVM.IAppContext
    ViewModel As TransferHistoryViewModel
    IsCancelled As Boolean
End Type

Private This As TView

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbClear_Click()
    If vbYes = MsgBox("Remove ALL saved tranfers?", vbExclamation + vbYesNo + vbDefaultButton2) Then
        This.ViewModel.Clear
        UpdateListView
    End If
End Sub

Private Sub cmbLoad_Click()
    Me.Hide
End Sub

Private Sub cmbRefresh_Click()
    This.ViewModel.Refresh
    UpdateListView
End Sub

Private Sub cmbRemoveWS_Click()
    If vbYes = MsgBox("Remove ENTIRE transfer history (including hidden worksheet)?", vbExclamation + vbYesNo + vbDefaultButton2) Then
        This.ViewModel.Remove
        MsgBox "Transfer Table History removed!", vbExclamation + vbOKOnly, "Transfer Table History" 'TODO Consts
        OnCancel
    End If
End Sub

Private Sub lvTransferInstructions_DblClick()
    If Not This.ViewModel.SelectedInstruction Is Nothing Then
        Me.Hide
    End If
End Sub

Private Sub lvTransferInstructions_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.TrySelect Item.key
    If Not This.ViewModel.SelectedInstruction Is Nothing Then
        Me.txtTransferInstruction.Value = This.ViewModel.SelectedInstruction.ToString
        Me.cmbLoad.Enabled = True
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function InitializeView() As Boolean
    InitializeView = True
    Me.cmbLoad.Enabled = False

    If This.ViewModel.HasHistory = False Then
        If vbYes = MsgBox("No history found. Initialize?", vbInformation + vbYesNo + vbDefaultButton1, "Transfer Table History") Then
            This.ViewModel.Create
        Else
            InitializeView = False
        End If
    End If
    
    This.ViewModel.Refresh
    This.ViewModel.InitializeListView Me.lvTransferInstructions
    UpdateListView
End Function

Private Function IView_ShowDialog(ByVal ViewModel As IViewModel) As Boolean
    Set This.ViewModel = ViewModel
    
    If InitializeView Then
        Me.Show
        IView_ShowDialog = Not This.IsCancelled
    Else
        OnCancel
    End If
End Function

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub UpdateListView()
    This.ViewModel.ItemsToListView Me.lvTransferInstructions
End Sub

