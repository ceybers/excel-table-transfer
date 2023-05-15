VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TestTableProperties")
Option Explicit

Public Enum ttWorkbookLocation
    Undefined
    LocalStorage
    RemoteStorage
    OneDriveStorage
    SharePointStorage
    UnsavedStorage
End Enum

Private Type TState
    Workbook As Workbook
End Type
Private This As TState

Property Get Workbook() As Workbook
    Set Workbook = This.Workbook
End Property

Public Property Set Workbook(ByVal vNewValue As Workbook)
    Debug.Assert TypeOf vNewValue Is Workbook
    Set This.Workbook = vNewValue
End Property
    
Public Property Get StorageLocation() As ttWorkbookLocation
    If This.Workbook.Path = vbNullString Then
        StorageLocation = UnsavedStorage
    ElseIf Left$(This.Workbook.Path, 2) = "\\" Then
        StorageLocation = RemoteStorage
    ElseIf Mid$(This.Workbook.Path, 2, 2) = ":\" Then
        StorageLocation = LocalStorage
        ' or onedrive, usually '%USERPROFILE%\OneDrive'
    ElseIf InStr(ThisWorkbook.Path, ".sharepoint.com/") > 0 Then
        StorageLocation = SharePointStorage
    End If
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = This.Workbook.ReadOnly
End Property

Public Property Get AutoSaveOn() As Boolean
    AutoSaveOn = This.Workbook.AutoSaveOn
End Property

Public Property Get ProtectStructure() As Boolean
    ProtectStructure = This.Workbook.ProtectStructure
End Property

