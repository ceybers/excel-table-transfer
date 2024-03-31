VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About Table Transfer Tool"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule HungarianNotation
'@Folder "MVVM.Views"
Option Explicit

Private Sub cboClose_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Me.lblHeader.Caption = APP_TITLE
    Me.lblVersion.Caption = APP_VERSION
    Me.lblCopyright.Caption = APP_COPYRIGHT
End Sub
