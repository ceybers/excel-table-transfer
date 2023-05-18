VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CityToComboBoxConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.Example2.Converters"
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    If TypeOf Source Is Scripting.Dictionary Then
        ConvertDictionary Source, Target
    Else
        ConvertDefault Source, Target
    End If
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding)
    Stop
End Sub

Private Sub ConvertDefault(ByVal SelectedItemKey As String, ByVal ComboBox As ComboBox)
    If SelectedItemKey = vbNullString Then
        Set ComboBox.Value = vbNullString
        Exit Sub
    End If
    
    ComboBox.Value = SelectedItemKey
End Sub

Private Sub ConvertDictionary(ByVal Cities As Scripting.Dictionary, ByVal ComboBox As ComboBox)
    Dim Key As Variant
    For Each Key In Cities
        Dim ThisCity As City
        Set ThisCity = Cities.Item(Key)
        Debug.Assert TypeOf ThisCity Is City
        
        ComboBox.AddItem ThisCity.Name
    Next Key
End Sub

Public Sub InitializeListView(ByVal ComboBox As ComboBox)
    With ComboBox
        .Clear
    End With
End Sub

