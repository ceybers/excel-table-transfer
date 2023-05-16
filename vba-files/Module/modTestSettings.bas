Attribute VB_Name = "modTestSettings"
'@Folder("VBAProject")
Option Explicit

Public Sub AAATEST()
    Dim s As SettingsModel
    Set s = SettingsModel.Create(ActiveWorkbook)
    's.Delete
    'Stop
    
    TestWorkbook s
    TestTable s
    TestCollection s
    
    s.DebugPrint
End Sub


Private Sub TestCollection(ByVal s As SettingsModel)
    Dim Coll As Collection
    Set Coll = New Collection
    Coll.Add Item:="Alpha"
    Coll.Add Item:="Bravo"
    Coll.Add Item:="Charlie"
    Coll.Add Item:="Delta"
    
    s.Workbook.SetCollection "Collection1", Coll
    
    Set Coll = Nothing
    
    Set Coll = s.Workbook.GetCollection("Collection1")
    
    Debug.Print "Coll count = "; Coll.Count
End Sub

Private Sub TestTable(ByVal s As SettingsModel)
    Dim t As ISomeSettings
    Set t = s.Tables.Item(1)
    
    Dim tt As SomeSettings
    
    Dim IsFoobar1 As Boolean
    IsFoobar1 = t.GetFlag("foobar2")
    Debug.Print "Foobar1 (before) = "; IsFoobar1
    
    IsFoobar1 = Not IsFoobar1
    
    t.SetFlag "foobar2", IsFoobar1
    
    IsFoobar1 = t.GetFlag("foobar2")
    Debug.Print "Foobar1 (after) = "; IsFoobar1
    
    Debug.Print "---"
    
    Dim BarFoo As String
    
    BarFoo = t.GetSetting("barfoo")
    Debug.Print "Barfoo (before) = "; BarFoo
    
    If Len(BarFoo) > 12 Then BarFoo = vbNullString
    BarFoo = BarFoo & " lorem"
    
    t.SetSetting "barfoo", BarFoo

    BarFoo = t.GetSetting("barfoo")
    Debug.Print "Barfoo (after) = "; BarFoo
    
    Debug.Print "---"
End Sub

Private Sub TestWorkbook(ByVal s As SettingsModel)
    Dim FooBar As Boolean
    
    FooBar = s.Workbook.GetFlag("foobar")
    Debug.Print "Foobar (before) = "; FooBar
    
    FooBar = Not FooBar
    
    s.Workbook.SetFlag "foobar", FooBar

    FooBar = s.Workbook.GetFlag("foobar")
    Debug.Print "Foobar (after) = "; FooBar
    
    Debug.Print "---"
    
    Dim BarFoo As String
    
    BarFoo = s.Workbook.GetSetting("barfoo")
    Debug.Print "Barfoo (before) = "; BarFoo
    
    If Len(BarFoo) > 12 Then BarFoo = vbNullString
    BarFoo = BarFoo & " lorem"
    
    s.Workbook.SetSetting "barfoo", BarFoo

    BarFoo = s.Workbook.GetSetting("barfoo")
    Debug.Print "Barfoo (after) = "; BarFoo
    
    Debug.Print "---"
End Sub
