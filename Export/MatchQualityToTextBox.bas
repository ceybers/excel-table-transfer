Attribute VB_Name = "MatchQualityToTextBox"
'@Folder "MVVM.ValueConverters"
Option Explicit

Public Sub Initialize(ByVal TextBox As MSForms.TextBox)
End Sub

Public Sub Load(ByVal TextBox As MSForms.TextBox, ByVal KeyColumnComparer As KeyColumnComparer)
    Dim MaxLines As Long
    With KeyColumnComparer
        MaxLines = .LeftOnly.Count
        If .Intersection.Count > MaxLines Then MaxLines = .Intersection.Count
        If .RightOnly.Count > MaxLines Then MaxLines = .RightOnly.Count
    End With
    
    Dim Table() As Variant
    ReDim Table(1 To MaxLines - 1)
    
    Dim i As Long
    For i = 1 To (MaxLines - 1)
        Dim Row() As Variant
        Row = Array(vbNullString, vbNullString, vbNullString)
        With KeyColumnComparer
            If (i < .LeftOnly.Count) Then Row(0) = .LeftOnly.Item(i)
            If (i < .Intersection.Count) Then Row(1) = .Intersection.Item(i)
            If (i < .RightOnly.Count) Then Row(2) = .RightOnly.Item(i)
        End With
        Table(i) = Join(Row, vbTab)
    Next i
    
    TextBox.Text = Join(Table, vbCrLf)
End Sub
