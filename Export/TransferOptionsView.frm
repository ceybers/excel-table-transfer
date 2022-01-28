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

Private Type TFrmKeyMapper2View
    ViewModel As TransferOptionsViewModel
    IsCancelled As Boolean
End Type

Private This As TFrmKeyMapper2View

Private Sub CheckBox1_Click()
    SetFlag ClearDestinationFirst, Me.CheckBox1
End Sub

Private Sub CheckBox2_Click()
    SetFlag TransferBlanks, Me.CheckBox2
End Sub

Private Sub CheckBox3_Click()
    SetFlag ReplaceEmptyOnly, Me.CheckBox3
End Sub

Private Sub CheckBox4_Click()
    SetFlag SourceFilteredOnly, Me.CheckBox4
End Sub

Private Sub CheckBox5_Click()
    SetFlag DestinationFilteredOnly, Me.CheckBox5
End Sub

Private Sub CheckBox6_Click()
    SetFlag appendunmapped, Me.CheckBox6
End Sub

Private Sub CheckBox7_Click()
    SetFlag RemoveUnmapped, Me.CheckBox7
End Sub

Private Sub CheckBox8_Click()
    SetFlag saveToHistory, Me.CheckBox8
End Sub

Private Sub SetFlag(ByVal flag As TransferOptionsEnum, ByRef cb As MSForms.CheckBox)
    This.ViewModel.Flags = modTestTransferOptions.SetFlag(This.ViewModel.Flags, flag, cb.value)
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
    This.IsCancelled = True
    Me.Hide
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As IViewModel) As Boolean
    Set This.ViewModel = ViewModel
    This.IsCancelled = False
    
    LoadFlags
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub LoadFlags()
    LoadFlag ClearDestinationFirst, Me.CheckBox1
    LoadFlag TransferBlanks, Me.CheckBox2
    LoadFlag ReplaceEmptyOnly, Me.CheckBox3
    LoadFlag SourceFilteredOnly, Me.CheckBox4
    LoadFlag DestinationFilteredOnly, Me.CheckBox5
    LoadFlag appendunmapped, Me.CheckBox6
    LoadFlag RemoveUnmapped, Me.CheckBox7
    LoadFlag saveToHistory, Me.CheckBox8
End Sub

Private Sub LoadFlag(ByVal flag As TransferOptionsEnum, ByVal cb As MSForms.CheckBox)
    cb.value = modTestTransferOptions.HasFlag(This.ViewModel.Flags, flag)
End Sub
