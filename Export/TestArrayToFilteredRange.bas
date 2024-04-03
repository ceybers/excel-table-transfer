Attribute VB_Name = "TestArrayToFilteredRange"
'@IgnoreModule
'@TestModule
'@Folder "Tests.Helpers"
Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Private Worksheet As Worksheet
    
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
    Set Worksheet = ThisWorkbook.Worksheets.Item("TestArrayToFilteredRange")
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

'@TestMethod("ArrayToFilteredRange")
Private Sub TestArrayToFilteredRange()
    On Error GoTo TestFail
    
    'Arrange:
    Worksheet.Rows.Item(2).Hidden = True
    Worksheet.Rows.Item(5).Hidden = True
    
    Dim SourceArray(1 To 6, 1 To 1) As Variant
    SourceArray(1, 1) = "a1"
    SourceArray(2, 1) = "a2"
    SourceArray(3, 1) = "a3"
    SourceArray(4, 1) = "a4"
    SourceArray(5, 1) = "a5"
    SourceArray(6, 1) = "a6"
    
    
    'Act:
    ArrayHelpers.ArrayToFilteredRange SourceArray, Worksheet.Range("A1:A6")
    
    Worksheet.Rows.Item(2).Hidden = False
    Worksheet.Rows.Item(5).Hidden = False
    
    'Assert:
    With Worksheet.Range("A1:A6").Cells
        Debug.Assert .Item(1).Value2 = "a1"
        Debug.Assert .Item(2).Value2 = vbNullString
        Debug.Assert .Item(3).Value2 = "a3"
        Debug.Assert .Item(4).Value2 = "a4"
        Debug.Assert .Item(5).Value2 = "hidden"
        Debug.Assert .Item(6).Value2 = "a6"
    End With
    
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
