Attribute VB_Name = "MatchQualityToComboBox"
'@Folder "MVVM.ValueConverters"
Option Explicit

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
    Dim Intersection As Long
    Intersection = Comparer.Intersection.Count
    
    Dim Total As Long
    Total = Comparer.LeftOnly.Count + Comparer.Intersection.Count + Comparer.RightOnly.Count
    
    If Total = 0 Then
        GetKeyColumnComparerString = MSG_ZERO_KEYS
    Else
    GetKeyColumnComparerString = Format$(Intersection / Total, "0%") & " (" & CStr(Intersection) & _
        "/" & CStr(Total) & " keys intersect)"
    End If
End Function


