VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Models"
Option Explicit

Private Const ROOT_NODE As String = "TestSettings"

Private Type TState
    Workbook As ISomeSettings
    Tables As Collection 'TableSettingsCollection
    CustomXMLPart As CustomXMLPart
End Type
Private This As TState

Public Property Get CustomXMLPart() As CustomXMLPart
    Set CustomXMLPart = CustomXMLPart
End Property

Public Property Set CustomXMLPart(ByVal vNewValue As CustomXMLPart)
    Set This.CustomXMLPart = vNewValue
End Property

Public Property Get Workbook() As ISomeSettings
    Set Workbook = This.Workbook
End Property

Public Property Set Workbook(ByVal vNewValue As ISomeSettings)
    Set This.Workbook = vNewValue
End Property

Public Property Get Tables() As Collection
    Set Tables = This.Tables
End Property

Public Function Create(ByVal Workbook As Workbook) As SettingsModel
    Set Create = New SettingsModel
    With Create
        Set .CustomXMLPart = GetCustomXMLPart(Workbook)
        Dim WorkbookSettings As SomeSettings
        Set WorkbookSettings = New SomeSettings
        WorkbookSettings.XPathPrefix = "/" & ROOT_NODE & "/Workbook"
        Set WorkbookSettings.CustomXMLPart = GetCustomXMLPart(Workbook)
        Set .Workbook = WorkbookSettings
        
        Dim TableSettings As SomeSettings
        Set TableSettings = New SomeSettings
        TableSettings.XPathPrefix = "/" & ROOT_NODE & "/Tables/Table[@Name='Table1']"
        Set TableSettings.CustomXMLPart = GetCustomXMLPart(Workbook)
        
        'Set This.Tables = New Collection
        .Tables.Add Item:=TableSettings, Key:="Table1"
    End With
End Function

Private Function GetCustomXMLPart(ByVal Workbook As Workbook) As CustomXMLPart
    Dim Result As CustomXMLNode
    Dim CustomXMLPart As CustomXMLPart
    For Each CustomXMLPart In Workbook.CustomXMLParts
        Set Result = CustomXMLPart.SelectSingleNode("/" & ROOT_NODE)
        If Not Result Is Nothing Then
            Set GetCustomXMLPart = CustomXMLPart
            Exit Function
        End If
    Next CustomXMLPart
    
    Workbook.CustomXMLParts.Add "<" & ROOT_NODE & "/>"
    Set GetCustomXMLPart = Workbook.CustomXMLParts.Item(ActiveWorkbook.CustomXMLParts.Count)
End Function

Public Sub Save()

End Sub

Public Sub Delete()
    GetCustomXMLPart(ActiveWorkbook).Delete
End Sub

Public Sub DebugPrint()
    'MsgBox This.CustomXMLPart.XML
    Debug.Print This.CustomXMLPart.XML
End Sub

Private Sub Class_Initialize()
    Set This.Tables = New Collection
End Sub
