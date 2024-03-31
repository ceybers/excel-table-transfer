Attribute VB_Name = "TablePickerToFrame"
'@Folder("MVVM2.ValueConverters")
Option Explicit

Private Const TAG_WORKBOOK As String = "WORKBOOK"
Private Const TAG_TABLE As String = "TABLE"
Private Const NO_TABLE_SELECTED As String = "(No table selected)"

Public Sub UpdateControls(ByVal ViewModel As TablePickerViewModel, ByVal Frame As Frame, ByVal Direction As TransferDirection)
    Dim Textbox As MSForms.Textbox
    
    If TryGetControlByTag(Frame.Controls, TAG_WORKBOOK, Textbox) Then
        If Direction = tdSource Then
            If Not ViewModel.SourceTable Is Nothing Then
                Textbox.Text = ViewModel.SourceTable.Parent.Parent.Name
                Textbox.Enabled = True
            Else
                Textbox.Text = NO_TABLE_SELECTED
                Textbox.Enabled = False
            End If
        Else
            If Not ViewModel.DestinationTable Is Nothing Then
                Textbox.Text = ViewModel.DestinationTable.Parent.Parent.Name
                Textbox.Enabled = True
            Else
                Textbox.Text = NO_TABLE_SELECTED
                Textbox.Enabled = False
            End If
        End If
    End If
    
    If TryGetControlByTag(Frame.Controls, TAG_TABLE, Textbox) Then
        If Direction = tdSource Then
            If Not ViewModel.SourceTable Is Nothing Then
                Textbox.Text = ViewModel.SourceTable.Name
                Textbox.Enabled = True
            Else
                Textbox.Text = NO_TABLE_SELECTED
                Textbox.Enabled = False
            End If
        Else
            If Not ViewModel.DestinationTable Is Nothing Then
                Textbox.Text = ViewModel.DestinationTable.Name
                Textbox.Enabled = True
            Else
                Textbox.Text = NO_TABLE_SELECTED
                Textbox.Enabled = False
            End If
        End If
    End If
End Sub

Public Function TryGetControlByTag(ByVal Controls As Controls, ByVal Tag As String, ByRef OutControl As Control) As Boolean
    Dim Control As Control
    For Each Control In Controls
        If Control.Tag = Tag Then
            Set OutControl = Control
            TryGetControlByTag = True
            Exit Function
        End If
    Next Control
End Function
