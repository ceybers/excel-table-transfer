Attribute VB_Name = "TestArrayToFilteredRange"
'@TestModule
'@Folder "Tests.Helpers"
Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Dim Worksheet As Worksheet
    
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
    Set Worksheet = ThisWorkbook.Worksheets("TestArrayToFilteredRange")
    With Worksheet
        .Cells.Clear
        .Range("A1").Value2 = "a"
        .Range("A2").Value2 = vbNullString
        .Range("A3").Value2 = "c"
        .Range("A4").Value2 = "d"
        .Range("A5").Value2 = "hidden"
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Worksheet.Cells.Clear
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod1()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Worksheet.rows(2).Hidden = True
    Worksheet.rows(5).Hidden = True
    
    Dim SourceArray(1 To 6, 1 To 1) As Variant
    SourceArray(1, 1) = "a1"
    SourceArray(2, 1) = "a2"
    SourceArray(3, 1) = "a3"
    SourceArray(4, 1) = "a4"
    SourceArray(5, 1) = "a5"
    SourceArray(6, 1) = "a6"
    
    ArrayHelpers.ArrayToFilteredRange SourceArray, Worksheet.Range("A1:A6")
    
    Worksheet.rows(2).Hidden = False
    Worksheet.rows(5).Hidden = False
    
    With Worksheet.Range("A1:A6").Cells
        Debug.Assert .Item(1).Value2 = "a1"
        Debug.Assert .Item(2).Value2 = vbNullString
        Debug.Assert .Item(3).Value2 = "a3"
        Debug.Assert .Item(4).Value2 = "a4"
        Debug.Assert .Item(5).Value2 = "hidden"
        Debug.Assert .Item(6).Value2 = "a6"
    End With
    
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
