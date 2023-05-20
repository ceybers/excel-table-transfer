VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyDocsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PersistentStorage.MyDocsSettings"
Option Explicit
Implements ISettings

Private Const DO_DEBUG As Boolean = False

Private Const SUBFOLDER As String = "Documents\Excel Macros"
Private Const Filename As String = "tabletransfertool.ini"
Private Const UUID As String = "{eb37a119-14e2-47c5-ab73-e46804ff84b5}"
Private Const FLAGS_HEADER As String = "FLAGS"
Private Const TRUE_CONST As String = "TRUE"
Private Const FALSE_CONST As String = "FALSE"

Private Type TState
    Lines() As String
    Flags As Scripting.Dictionary
    AutoSave As Boolean
End Type
Private This As TState

Private Function GetFolder() As String
    GetFolder = Environ$("USERPROFILE") & "\" & SUBFOLDER
End Function

Private Function GetFullPath() As String
    GetFullPath = Environ$("USERPROFILE") & "\" & SUBFOLDER & "\" & Filename
End Function

Private Sub LoadFromFile()
    Open GetFullPath() For Binary As #1
    
    Dim FullData As String
    '@Ignore AssignmentNotUsed
    FullData = Space$(LOF(1))
    Get #1, , FullData
    Close #1
    
    This.Lines = Split(FullData, vbCrLf)
    
    If FullData = vbNullString Then
        SetDefaultFlags
    End If
End Sub

Private Sub SaveData()
    Open GetFullPath() For Output As #1
    Print #1, UUID
    Print #1, FLAGS_HEADER
    
    Dim FlagKey As Variant
    For Each FlagKey In This.Flags.Keys
    Print #1, vbTab; FlagKey; vbTab; This.Flags.Item(FlagKey)
    Next FlagKey
    Close #1
End Sub

Private Sub CheckIfEmpty()
    Dim FileSystemObject As Object
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    If Not FileSystemObject.FolderExists(GetFolder()) Then
        FileSystemObject.CreateFolder GetFolder()
    End If
    
    If Not FileSystemObject.FileExists(GetFullPath()) Then
        FileSystemObject.CreateTextFile GetFullPath()
        SetDefaultFlags
        SaveData
    End If
    
    Set FileSystemObject = Nothing
End Sub

Private Sub Class_Initialize()
    Set This.Flags = New Scripting.Dictionary
    This.AutoSave = True
End Sub

Private Function IsMySettings() As Boolean
    If UBound(This.Lines) = -1 Then Exit Function
    IsMySettings = (This.Lines(0) = UUID)
End Function

Private Sub LoadFlags()
    Dim i As Long
    Dim ThisLine() As String
    Dim Extracting As Boolean
    
    If DO_DEBUG Then Debug.Print "Loaded "; UBound(This.Lines); " lines"
    
    For i = 1 To UBound(This.Lines)
        ThisLine = Split(This.Lines(i), vbTab)
        If Extracting Then
            If UBound(ThisLine) = 2 Then
                This.Flags.Add Key:=ThisLine(1), Item:=ThisLine(2)
            Else
                Extracting = False
            End If
        Else
            If ThisLine(0) = FLAGS_HEADER Then Extracting = True
        End If
    Next i
End Sub

Private Function ISettings_GetFlag(ByVal FlagName As String) As Boolean
    ISettings_GetFlag = (UCase$(This.Flags.Item(FlagName)) = TRUE_CONST)
    If DO_DEBUG Then Debug.Print "ISettings_GetFlag("; FlagName; ") = "; ISettings_GetFlag
End Function

Private Sub SetDefaultFlags()
    Set This.Flags = New Scripting.Dictionary
End Sub

Private Sub ISettings_SetFlag(ByVal FlagName As String, ByVal FlagValue As Boolean)
    If DO_DEBUG Then Debug.Print "ISettings_SetFlag("; FlagName; ") = "; FlagValue
    
    If FlagValue = True Then
        This.Flags.Item(FlagName) = TRUE_CONST
    Else
        This.Flags.Item(FlagName) = FALSE_CONST
    End If
    
    If This.AutoSave Then SaveData
End Sub

Public Sub Load()
    CheckIfEmpty
    LoadFromFile
    If Not IsMySettings Then Exit Sub
    LoadFlags
End Sub

Public Sub Save()
    SaveData
End Sub

Private Sub ISettings_Reset()
    SetDefaultFlags
    SaveData
End Sub

Private Function ISettings_GetSetting(ByVal SettingName As String) As String
    Err.Raise Number:=42, Description:="NYI ISettings_GetSetting"
End Function

Private Sub ISettings_SetSetting(ByVal SettingName As String, ByVal Value As String)
    Err.Raise Number:=42, Description:="NYI ISettings_SetSetting"
End Sub

Private Function ISettings_GetCollection(ByVal CollectionName As String) As Collection
    Err.Raise Number:=42, Description:="NYI ISettings_GetSetting"
End Function

Private Sub ISettings_SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)
    Err.Raise Number:=42, Description:="NYI ISettings_SetSetting"
End Sub
