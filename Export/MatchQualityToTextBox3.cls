VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchQualityToTextBox3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.ValueConverters")
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object) ' read
    Load Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding) ' write
    ' Read-only
End Sub

Public Sub Load(ByVal TextBox As MSForms.TextBox, ByVal KeyColumn As KeyColumn)
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

