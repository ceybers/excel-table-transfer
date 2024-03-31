Attribute VB_Name = "KeyColumnsToComboBox"
'@Folder("MVVM2.ValueConverters")
Option Explicit

Private Const NO_COLUMN_SELECTED As String = "(No column selected)"
Private Const NO_COLUMN_COLOR As Long = 8421504 ' RGB(128,128,128)

Public Sub Initialize(ByVal ComboBox As MSForms.ComboBox)
    'ComboBox.Enabled = True
    'ComboBox.Locked = True
End Sub

Public Sub Load(ByVal ComboBox As MSForms.ComboBox, ByVal KeyColumns As KeyColumns)
    Debug.Assert Not KeyColumns Is Nothing
    
    If KeyColumns.Selected Is Nothing Then
        ComboBox.Text = NO_COLUMN_SELECTED
        ComboBox.Enabled = False
    Else
        ComboBox.Text = KeyColumns.Selected.Count & " of " & KeyColumns.Selected.Range.Cells.Count & " distinct"
        ComboBox.Enabled = True
    End If
End Sub
