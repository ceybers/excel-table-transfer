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

'@IgnoreModule HungarianNotation, SetAssignmentWithIncompatibleObjectType
'@Folder "MVVM2.Views"
Option Explicit

Private Sub cmbClose_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Me.lblHeader.Caption = "Table Transfer Tool"
    Me.lblVersion.Caption = "Version 1.8.1-dev"
    Me.lblCopyright.Caption = "Copyright © 2024 Craig Eybers" & vbCrLf & "All rights reserved."
End Sub
