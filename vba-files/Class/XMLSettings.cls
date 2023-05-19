VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PersistentSettings.Settings"
Option Explicit
Implements ISettings

Private Const FIRST_CHILD As String = "TableTransferTool"
Private Const DEFAULT_XML As String = "<TableTransferTool><Flags/><Settings/><History/></TableTransferTool>"

Private Type TState
    CustomXMLParts As CustomXMLParts
    CustomXMLPart As CustomXMLPart
    XMLDoc As MSXML2.DOMDocument60
End Type
Private this As TState

Private Sub ISettings_Load(Optional ByVal mWorkbook As Object)
    Dim Workbook As Workbook
    Set Workbook = mWorkbook
    Set this.CustomXMLParts = Workbook.CustomXMLParts
    
    If Not AlreadyExistsInWorkbook Then AddToWorkbook
    
    GetMyXMLDoc
End Sub

Private Sub ISettings_Reset()
    If this.CustomXMLPart Is Nothing Then Exit Sub
    this.CustomXMLPart.Delete
End Sub

Private Sub ISettings_Save()
    'NOOP
End Sub

Private Sub AddToWorkbook()
    this.CustomXMLParts.Add XML:=DEFAULT_XML
End Sub

Private Function StringToXML(ByVal XMLString As String) As MSXML2.DOMDocument60
    Set StringToXML = New MSXML2.DOMDocument60
    StringToXML.LoadXML XMLString
End Function

Private Function IsMyXMLDoc(ByVal CustomXMLPart As CustomXMLPart) As Boolean
    IsMyXMLDoc = (StringToXML(CustomXMLPart.XML).FirstChild.BaseName = FIRST_CHILD)
End Function

Private Function AlreadyExistsInWorkbook() As Boolean
   Dim CustomXMLPart As CustomXMLPart
   For Each CustomXMLPart In this.CustomXMLParts
        If IsMyXMLDoc(CustomXMLPart) Then
            AlreadyExistsInWorkbook = True
            Exit Function
        End If
    Next CustomXMLPart
End Function

Private Function GetMyXMLDoc() As Boolean
    Dim CustomXMLPart As CustomXMLPart
    For Each CustomXMLPart In this.CustomXMLParts
        If IsMyXMLDoc(CustomXMLPart) Then
            Set this.CustomXMLPart = CustomXMLPart
            Set this.XMLDoc = StringToXML(CustomXMLPart.XML)
            GetMyXMLDoc = True
            Exit Function
        End If
    Next CustomXMLPart
End Function

Public Function HasFlag(ByVal FlagName As String) As Boolean
    Dim result As CustomXMLNode
    Set result = this.CustomXMLPart.SelectSingleNode("//TableTransferTool/Flags/" & FlagName)
    HasFlag = Not result Is Nothing
End Function

Private Sub ISettings_SetFlag(ByVal FlagName As String, ByVal FlagValue As Boolean)
    Dim result As CustomXMLNode
    Set result = this.CustomXMLPart.SelectSingleNode("//TableTransferTool/Flags/" & FlagName)
    If result Is Nothing Then
        Dim FlagNode As CustomXMLNode
        Set FlagNode = this.CustomXMLPart.SelectSingleNode("//TableTransferTool/Flags")
        FlagNode.AppendChildNode Name:=FlagName, NodeValue:=CStr(FlagValue)
    Else
        result.text = CStr(FlagValue)
    End If
End Sub
 
Private Function ISettings_GetFlag(ByVal FlagName As String) As Boolean
    ' TODO Consider Implement TryGetFlag(ByVal FlagName, ByRef Result) as Boolean
    If Not HasFlag(FlagName) Then Exit Function
    
    Dim result As CustomXMLNode
    Set result = this.CustomXMLPart.SelectSingleNode("//TableTransferTool/Flags/" & FlagName)
    ISettings_GetFlag = CBool(result.text)
End Function

Public Function HasSetting(ByVal SettingName As String) As Boolean
    Dim result As CustomXMLNode
    Set result = this.CustomXMLPart.SelectSingleNode("//TableTransferTool/Settings/" & SettingName)
    HasSetting = Not result Is Nothing
End Function

Public Sub SetSetting(ByVal SettingName As String, ByVal SettingValue As String)
    Dim result As CustomXMLNode
    Set result = this.CustomXMLPart.SelectSingleNode("//TableTransferTool/Settings/" & SettingName)
    If result Is Nothing Then
        Dim SettingNode As CustomXMLNode
        Set SettingNode = this.CustomXMLPart.SelectSingleNode("//TableTransferTool/Settings")
        SettingNode.AppendChildNode Name:=SettingName, NodeValue:=SettingValue
    Else
        result.text = SettingValue
    End If
End Sub

Public Function GetSetting(ByVal SettingName As String) As String
    If Not HasSetting(SettingName) Then Exit Function
    
    Dim result As CustomXMLNode
    Set result = this.CustomXMLPart.SelectSingleNode("//TableTransferTool/Settings/" & SettingName)
    GetSetting = result.text
End Function
