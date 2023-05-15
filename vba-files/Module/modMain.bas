Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

' Reference: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms756987(v=vs.85)
' Reference: https://learn.microsoft.com/en-us/office/vba/api/office.customxmlnode

'@EntryPoint "DoLoad"
Public Sub DoLoad()
    Dim XMLSettings As XMLSettings
    Set XMLSettings = New XMLSettings
    XMLSettings.Load ThisWorkbook
    
    Dim mName As String
    mName = ActiveSheet.Range("txtName").Value2
    
    Dim mValue As String
    mValue = XMLSettings.GetSetting(mName)
    
    ActiveSheet.Range("txtValue").Value2 = mValue
End Sub

'@EntryPoint "DoSave"
Public Sub DoSave()
    Dim XMLSettings As XMLSettings
    Set XMLSettings = New XMLSettings
    XMLSettings.Load ThisWorkbook

    Dim mName As String
    mName = ActiveSheet.Range("txtName").Value2
    
    Dim mValue As String
    mValue = ActiveSheet.Range("txtValue").Value2
    
    XMLSettings.SetSetting mName, mValue
End Sub

'@EntryPoint "DoPrint"
Public Sub DoPrint()
    Dim CustomXMLPart As CustomXMLPart
    For Each CustomXMLPart In ThisWorkbook.CustomXMLParts
        PrintCustomXMLPart CustomXMLPart
    Next CustomXMLPart
End Sub

'@EntryPoint "DoReset"
Public Sub DoReset()
    Dim XMLSettings As XMLSettings
    Set XMLSettings = New XMLSettings
    XMLSettings.Load ThisWorkbook
    XMLSettings.Reset
End Sub

Private Sub PrintCustomXMLPart(ByVal CustomXMLPart As CustomXMLPart)
    Debug.Print "CustomXMLPart.ID: "; CustomXMLPart.ID
    Debug.Print "CustomXMLPart.NamespaceURI: "; CustomXMLPart.NamespaceURI
    Debug.Print "CustomXMLPart.XML: "; CustomXMLPart.XML
    Debug.Print " "
End Sub
