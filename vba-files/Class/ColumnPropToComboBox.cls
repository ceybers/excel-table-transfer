VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPropToComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TestTableProperties")
Option Explicit
Implements IValueConverter

Private Const DEFAULT_ITEM As String = "(undefined)"

Private Sub IValueConverter_Convert(ByVal Control As Object, ByVal Object As Object)
    Debug.Assert TypeOf Control Is ComboBox
    Debug.Assert TypeOf Object Is ColumnProperty
    Convert Control, Object
End Sub

Private Sub Convert(ByVal ComboBox As ComboBox, ByVal ColumnProperty As ColumnProperty)
    With ComboBox
        If ColumnProperty.ColumnType <> "Text" Then
            .AddItem ColumnProperty.Name & " (non-text)"
        ElseIf ColumnProperty.Unique = False Then
            .AddItem ColumnProperty.Name & " (non-unique)"
        Else
            .AddItem ColumnProperty.Name
        End If
    End With
End Sub

Private Sub IValueConverter_Initialize(ByVal Control As Object)
    Debug.Assert TypeOf Control Is ComboBox
    InitializeComboBox Control
End Sub

Private Sub InitializeComboBox(ByVal ComboBox As ComboBox)
    With ComboBox
        .Clear
        .AddItem DEFAULT_ITEM
        .Value = DEFAULT_ITEM
    End With
End Sub

