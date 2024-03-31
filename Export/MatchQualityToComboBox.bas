Attribute VB_Name = "MatchQualityToComboBox"
'@Folder("MVVM2.ValueConverters")
Option Explicit

Private Const NO_TWO_COLUMNS_SELECTED As String = vbNullString

Public Sub Initialize(ByVal ComboBox As MSForms.ComboBox)
    'ComboBox.Locked = True
End Sub

Public Sub Load(ByVal ComboBox As MSForms.ComboBox, ByVal KeyColumnComparer As KeyColumnComparer)
    If KeyColumnComparer Is Nothing Then
        ComboBox.Text = NO_TWO_COLUMNS_SELECTED
        ComboBox.Enabled = False
    Else
        ComboBox.Text = GetKeyColumnComparerString(KeyColumnComparer)
        ComboBox.Enabled = True
    End If
End Sub

Private Function GetKeyColumnComparerString(ByVal Comparer As KeyColumnComparer) As String
    GetKeyColumnComparerString = CStr(Comparer.Intersection.Count) & " of " & _
        CStr(Comparer.LeftOnly.Count + Comparer.Intersection.Count + Comparer.RightOnly.Count) & " matches"
End Function


