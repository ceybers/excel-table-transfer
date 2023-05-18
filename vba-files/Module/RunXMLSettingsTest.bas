Attribute VB_Name = "RunXMLSettingsTest"
'@Folder "SomeSettingsModel"
Option Explicit

'@EntryPoint "XMLSettingsTest"
Public Sub XMLSettingsTest()
    Dim TestModel As ISettingsModel
    Set TestModel = XMLSettingsModel.Create(ActiveWorkbook, "TestSettings")
    'TestModel.Reset
    
    TestWorkbook TestModel
    TestTable TestModel
    TestCollection TestModel
    
    'TestModel.DebugPrint
    Debug.Print "END"
    Debug.Print "---"
End Sub

'@EntryPoint "XMLSettingsReset"
Public Sub XMLSettingsReset()
    Dim s As XMLSettingsModel
    Set s = XMLSettingsModel.Create(ActiveWorkbook, "TestSettings")
    s.Delete
End Sub

Private Sub TestCollection(ByVal s As ISettingsModel)
    Debug.Print "Testing Collection Function"
    
    Dim Coll As Collection
    Set Coll = New Collection
    Coll.Add Item:="Alpha"
    Coll.Add Item:="Bravo"
    Coll.Add Item:="Charlie"
    Coll.Add Item:="Delta"
    
    Debug.Print " Setting Collection1..."
    s.Workbook.SetCollection "Collection1", Coll
    
    Set Coll = Nothing
    
    Debug.Print " Getting Collection1..."
    Set Coll = s.Workbook.GetCollection("Collection1")
    
    Debug.Print " Coll count = "; Coll.Count
    
    Debug.Print "---"
End Sub

Private Sub TestTable(ByVal s As ISettingsModel)
    Debug.Print "Testing Table-level Functions"
    Dim t As ISettings
    Set t = s.Table("Table9")
    
    Dim tt As XMLSettings
    
    Dim IsFoobar1 As Boolean
    IsFoobar1 = t.GetFlag("foobar2")
    Debug.Print " Foobar1 (before) = "; IsFoobar1
    
    IsFoobar1 = Not IsFoobar1
    
    t.SetFlag "foobar2", IsFoobar1
    
    IsFoobar1 = t.GetFlag("foobar2")
    Debug.Print " Foobar1 (after) = "; IsFoobar1
    
    Debug.Print vbNullString
    
    Dim BarFoo As String
    
    BarFoo = t.GetSetting("barfoo")
    Debug.Print " Barfoo (before) = "; BarFoo
    
    If Len(BarFoo) > 12 Then BarFoo = vbNullString
    BarFoo = BarFoo & " lorem"
    
    t.SetSetting "barfoo", BarFoo

    BarFoo = t.GetSetting("barfoo")
    Debug.Print " Barfoo (after) = "; BarFoo
    
    Debug.Print "---"
End Sub

Private Sub TestWorkbook(ByVal s As ISettingsModel)
    Debug.Print "Testing Workbook-level Functions"
    Dim FooBar As Boolean
    
    FooBar = s.Workbook.GetFlag("foobar")
    Debug.Print " Foobar (before) = "; FooBar
    
    FooBar = Not FooBar
    
    s.Workbook.SetFlag "foobar", FooBar

    FooBar = s.Workbook.GetFlag("foobar")
    Debug.Print " Foobar (after) = "; FooBar
    
    Debug.Print vbNullString
    
    Dim BarFoo As String
    
    BarFoo = s.Workbook.GetSetting("barfoo")
    Debug.Print " Barfoo (before) = "; BarFoo
    
    If Len(BarFoo) > 12 Then BarFoo = vbNullString
    BarFoo = BarFoo & " lorem"
    
    s.Workbook.SetSetting "barfoo", BarFoo

    BarFoo = s.Workbook.GetSetting("barfoo")
    Debug.Print " Barfoo (after) = "; BarFoo
    
    Debug.Print "---"
End Sub
