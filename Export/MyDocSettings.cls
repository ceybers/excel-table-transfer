VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyDocSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "PersistentStorage.MyDocSettings"
Option Explicit
Implements ISettings

Private Const DO_DEBUG As Boolean = False

Private Const SUBFOLDER As String = "Documents\Excel Macros"
Private Const FLAGS_HEADER As String = "[flags]"
Private Const SETTINGS_HEADER As String = "[settings]"
Private Const COLLECTIONS_HEADER As String = "[collection."
Private Const TRUE_CONST As String = "TRUE"
Private Const FALSE_CONST As String = "FALSE"
Private Const DELIMITER_CHAR As String = "="
Private Const COMMENT_CHAR As String = ";"

Private Type TState
    UUID As String
    Filename As String
    
    Lines() As String
    
    Flags As Scripting.Dictionary
    Settings As Scripting.Dictionary
    Collections As Scripting.Dictionary
    
    AutoSave As Boolean
End Type
Private This As TState

Public Property Get UUID() As String
    UUID = This.UUID
End Property

Public Property Let UUID(ByVal vNewValue As String)
    Debug.Assert Len(vNewValue) = 38
    Debug.Assert Left$(vNewValue, 1) = "{"
    Debug.Assert Right$(vNewValue, 1) = "}"
    
    This.UUID = vNewValue
End Property

Public Property Get Filename() As String
    Filename = This.Filename
End Property

Public Property Let Filename(ByVal vNewValue As String)
    This.Filename = vNewValue
End Property

' ---
Private Function ISettings_GetFlag(ByVal FlagName As String) As Boolean
    If Not IsKeyValid(FlagName) Then
        Err.Raise vbObjectError + 13, "MyDocSettings", "Invalid key"
    End If
    
    ISettings_GetFlag = (UCase$(This.Flags.Item(FlagName)) = TRUE_CONST)
End Function

Private Sub ISettings_SetFlag(ByVal FlagName As String, ByVal FlagValue As Boolean)
    If Not IsKeyValid(FlagName) Then
        Err.Raise vbObjectError + 13, "MyDocSettings", "Invalid key"
    End If
    
    If FlagValue = True Then
        This.Flags.Item(FlagName) = TRUE_CONST
    Else
        This.Flags.Item(FlagName) = FALSE_CONST
    End If
    
    If This.AutoSave Then SaveData
End Sub

Private Function ISettings_GetSetting(ByVal SettingName As String) As String
    If Not IsKeyValid(SettingName) Then Exit Function
    If Not This.Settings.Exists(SettingName) Then Exit Function
    ISettings_GetSetting = This.Settings.Item(SettingName)
    If DO_DEBUG Then Debug.Print "ISettings_GetSetting("; SettingName; ") = "; ISettings_GetSetting
End Function

Private Sub ISettings_SetSetting(ByVal SettingName As String, ByVal Value As String)
    If Not IsKeyValid(SettingName) Then Exit Sub
    
    If DO_DEBUG Then Debug.Print "ISettings_SetSetting("; SettingName; ") = "; Value
    
    This.Settings.Item(SettingName) = Value
    
    If This.AutoSave Then SaveData
End Sub

Private Function ISettings_GetCollection(ByVal CollectionName As String) As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    If Not This.Collections.Exists(CollectionName) Then
        'Err.Raise vbObjectError + 9, "MyDocSettings", "Collection does not exist"
        Set ISettings_GetCollection = Result
        Exit Function
    End If
    
    Dim Arr As Variant
    Arr = This.Collections.Item(CollectionName).Items
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Result.Add Arr(i)
    Next i
    
    Set ISettings_GetCollection = Result
End Function

Private Function ISettings_HasCollection(ByVal CollectionName As String) As Boolean
    ISettings_HasCollection = This.Collections.Exists(CollectionName)
End Function

Private Sub ISettings_SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)
    If Not IsKeyValid(CollectionName) Then
        Err.Raise vbObjectError + 9, "MyDocSettings", "Collection name is not valid."
    End If
    
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    
    Dim i As Long
    For i = 1 To Collection.Count
        Result.Add Key:=("N" & (i - 1)), Item:=CStr(Collection.Item(i))
    Next i
    
    Set This.Collections.Item(CollectionName) = Result
    
    If This.AutoSave Then SaveData
End Sub

Private Sub ISettings_Reset()
    This.Flags.RemoveAll
    This.Settings.RemoveAll
End Sub

' ---
Private Function GetFolder() As String
    GetFolder = Environ$("USERPROFILE") & "\" & SUBFOLDER
End Function

Private Function GetFullPath() As String
    GetFullPath = Environ$("USERPROFILE") & "\" & SUBFOLDER & "\" & This.Filename
End Function

Private Function CollectionNameToSection(ByVal CollectionName As String) As String
    CollectionNameToSection = COLLECTIONS_HEADER & CollectionName & "]"
End Function

Private Function IsMySettings() As Boolean
    If UBound(This.Lines) = -1 Then Exit Function
    IsMySettings = (This.Lines(0) = (COMMENT_CHAR & " " & This.UUID))
End Function
' ---

Private Sub Class_Initialize()
    Set This.Flags = New Scripting.Dictionary
    Set This.Settings = New Scripting.Dictionary
    Set This.Collections = New Scripting.Dictionary
    
    This.AutoSave = True
End Sub

Public Function Create(ByVal UUID As String, ByVal Filename As String) As ISettings
    Dim Result As MyDocSettings
    Set Result = New MyDocSettings
    With Result
        .UUID = UUID
        .Filename = Filename
        .Load
    End With
    
    Set Create = Result
End Function

Public Sub Load()
    If DO_DEBUG Then Debug.Print "MyDocSettings.Load()"
    
    CheckIfFileExists
    LoadFromFile
    If Not IsMySettings Then Exit Sub
    
    LoadSection FLAGS_HEADER, This.Flags
    LoadSection SETTINGS_HEADER, This.Settings
    
    Dim CollectionName As Variant
    For Each CollectionName In GetCollectionHeaders
        LoadCollection CollectionName
    Next CollectionName
End Sub

Private Sub LoadFromFile()
    Open GetFullPath() For Binary As #1
    
    Dim FullData As String
    '@Ignore AssignmentNotUsed
    FullData = Space$(LOF(1))
    Get #1, , FullData
    Close #1
    
    This.Lines = Split(FullData, vbCrLf)
End Sub

Private Sub LoadSection(ByVal SectionHeading As String, ByVal Dictionary As Scripting.Dictionary)
    If DO_DEBUG Then Debug.Print "Loaded "; UBound(This.Lines); " lines"
    
    Dim SectionStart As Long
    For SectionStart = 1 To UBound(This.Lines)
        If This.Lines(SectionStart) = SectionHeading Then Exit For
    Next SectionStart
    If SectionStart = UBound(This.Lines) Then Exit Sub
    
    Dim i As Long
    For i = (SectionStart + 1) To UBound(This.Lines)
        Dim ThisLine As String
        ThisLine = Trim$(This.Lines(i))
        
        Dim DelimiterIndex As Long
        DelimiterIndex = InStr(ThisLine, DELIMITER_CHAR)
        
        If Left$(ThisLine, 1) = COMMENT_CHAR Then DelimiterIndex = 0
        If Left$(ThisLine, 1) = "[" Then Exit Sub
        
        If DelimiterIndex > 0 Then
            Dim Key As String
            Dim Value As String
            Key = Trim$(Left$(ThisLine, DelimiterIndex - 1))
            Value = Trim$(Mid$(ThisLine, DelimiterIndex + 1))
            
            Dictionary.Add Key:=Key, Item:=Value
        End If
    Next i
End Sub

Private Sub LoadCollection(ByVal CollectionName As String)
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    This.Collections.Add Key:=CollectionName, Item:=Result
    
    Dim SectionName As String
    SectionName = CollectionNameToSection(CollectionName)
    
    LoadSection SectionName, Result
End Sub

Private Sub SaveData()
    Dim FileNumber As Long
    FileNumber = FreeFile
    
    Open GetFullPath() For Output As #FileNumber
    Print #FileNumber, COMMENT_CHAR; " "; UUID
    
    SaveSection FileNumber, FLAGS_HEADER, This.Flags
    SaveSection FileNumber, SETTINGS_HEADER, This.Settings
    
    Dim Key As Variant
    For Each Key In This.Collections.Keys
        SaveSection FileNumber, CollectionNameToSection(Key), This.Collections.Item(Key)
    Next
    
    Close #FileNumber
End Sub

Private Sub SaveSection(ByVal FileNumber As Long, ByVal SectionHeading As String, ByVal Dictionary As Scripting.Dictionary)
    Print #FileNumber, SectionHeading
    
    Dim Key As Variant
    For Each Key In Dictionary.Keys
        Print #FileNumber, Key; DELIMITER_CHAR; Dictionary.Item(Key)
    Next Key
End Sub

Private Function IsKeyValid(ByVal Key As String) As Boolean
    If Key = vbNullString Then Exit Function
    If InStr(Key, "=") > 0 Then Exit Function
    If InStr(Key, ";") > 0 Then Exit Function
    If InStr(Key, vbCr) > 0 Then Exit Function
    If InStr(Key, vbLf) > 0 Then Exit Function
    If Key <> Trim$(Key) Then Exit Function
    
    IsKeyValid = True
End Function

Private Function GetCollectionHeaders() As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim ThisLine As String
    Dim CollectionName As String
    Dim Index As Long
    Dim i As Long
    For i = 1 To UBound(This.Lines)
        ThisLine = Trim$(This.Lines(i))
        Index = InStr(ThisLine, COLLECTIONS_HEADER)
        If Index > 0 Then
            CollectionName = Mid$(ThisLine, Index + Len(COLLECTIONS_HEADER))
            CollectionName = Left$(CollectionName, Len(CollectionName) - 1)
            Result.Add Item:=CollectionName
        End If
    Next i
    
    Set GetCollectionHeaders = Result
End Function

Private Sub CheckIfFileExists()
    Dim FileSystemObject As Object
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    If Not FileSystemObject.FolderExists(GetFolder()) Then
        FileSystemObject.CreateFolder GetFolder()
    End If
    
    If Not FileSystemObject.FileExists(GetFullPath()) Then
        FileSystemObject.CreateTextFile GetFullPath()
    End If
    
    Set FileSystemObject = Nothing
End Sub
