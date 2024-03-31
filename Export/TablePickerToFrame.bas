Attribute VB_Name = "TablePickerToFrame"
'@Folder("MVVM2.ValueConverters")
Option Explicit

Private Const TAG_WORKBOOK As String = "WORKBOOK"
Private Const TAG_TABLE As String = "TABLE"
Private Const NO_TABLE_SELECTED As String = "(No table selected)"

Public Sub UpdateControls(ByVal ViewModel As TablePickerViewModel, ByVal Frame As Frame, ByVal Direction As TransferDirection)
    Dim TextBox As MSForms.TextBox
    
    If TryGetControlByTag(Frame.Controls, TAG_WORKBOOK, TextBox) Then
        If Direction = tdSource Then
            If Not ViewModel.SourceTable Is Nothing Then
                TextBox.Text = ViewModel.SourceTable.Parent.Parent.Name
                TextBox.Enabled = True
            Else
                TextBox.Text = NO_TABLE_SELECTED
                TextBox.Enabled = False
            End If
        Else
            If Not ViewModel.DestinationTable Is Nothing Then
                TextBox.Text = ViewModel.DestinationTable.Parent.Parent.Name
                TextBox.Enabled = True
            Else
                TextBox.Text = NO_TABLE_SELECTED
                TextBox.Enabled = False
            End If
        End If
    End If
    
    If TryGetControlByTag(Frame.Controls, TAG_TABLE, TextBox) Then
        If Direction = tdSource Then
            If Not ViewModel.SourceTable Is Nothing Then
                TextBox.Text = ViewModel.SourceTable.Name
                TextBox.Enabled = True
            Else
                TextBox.Text = NO_TABLE_SELECTED
                TextBox.Enabled = False
            End If
        Else
            If Not ViewModel.DestinationTable Is Nothing Then
                TextBox.Text = ViewModel.DestinationTable.Name
                TextBox.Enabled = True
            Else
                TextBox.Text = NO_TABLE_SELECTED
                TextBox.Enabled = False
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
