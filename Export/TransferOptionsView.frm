VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferOptionsView 
   Caption         =   "Transfer Options"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   OleObjectBlob   =   "TransferOptionsView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransferOptionsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TransferOptions")
Option Explicit
Implements IView

Public vm As IViewModel
Public flags As Integer

Private Type TFrmKeyMapper2View
    IsCancelled As Boolean
End Type

Private this As TFrmKeyMapper2View

Private Sub CheckBox1_Click()
    flags = modTestTransferOptions.SetFlag(flags, TransferOptionsEnum.ClearDestinationFirst, Me.CheckBox1.value)
End Sub

Private Sub CheckBox2_Click()
    flags = modTestTransferOptions.SetFlag(flags, TransferOptionsEnum.TransferBlanks, Me.CheckBox2.value)
End Sub

Private Sub CheckBox3_Click()
    flags = modTestTransferOptions.SetFlag(flags, TransferOptionsEnum.ReplaceEmptyOnly, Me.CheckBox3.value)
End Sub

' ---
Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbOK_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Function IView_ShowDialog(ByVal viewModel As IViewModel) As Boolean
    Set vm = viewModel
    this.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not this.IsCancelled
End Function
