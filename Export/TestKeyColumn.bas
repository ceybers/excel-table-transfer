Attribute VB_Name = "TestKeyColumn"
'@IgnoreModule
'@TestModule
'@Folder "Tests.Model"
Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("KeyColumn")
Private Sub TestKeyColumn()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Dim Worksheet As Worksheet
    Set Worksheet = ThisWorkbook.Worksheets.Item(2)
    
    Dim Range As Range
    Set Range = Worksheet.Range("A2:A5,A14")
    
    Dim Key As KeyColumn
    Set Key = KeyColumn.FromRange(Range, True, True)

    'DebugPrint Key
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub DebugPrint(ByVal Key As KeyColumn)
    Debug.Print "TEST KeyColumn"
    Debug.Print "===="
    Debug.Print "Distinct = " & Key.Count
    Debug.Print "Unique = " & Key.UniqueKeys.Count
    Debug.Print "IsDistinct = " & Key.IsDistinct
    Debug.Print "Errors = " & Key.ErrorCount
    Debug.Print "Blanks = " & Key.BlankCount
    Debug.Print "Find 'def' = " & Key.Find("def")
    Debug.Print "Find '1234567890' = " & Key.Find("1234567890")
    Debug.Print "Find 'Right Only2' = " & Key.Find("Right Only2")
    Debug.Print vbNullString
End Sub
