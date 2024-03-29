Attribute VB_Name = "TestColumnPairs"
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

'@TestMethod("ColumnPairs")
Private Sub TestColumnPairs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LHS As ListObject
    Dim RHS As ListObject
    
    Set LHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    Set RHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
    
    Dim colPairs As ColumnPairs
    Set colPairs = New ColumnPairs

    'Act:
    Dim ColPair As ColumnPair
    
    Set ColPair = ColumnPair.Create(LHS.ListColumns.Item(2), RHS.ListColumns.Item(2))
    colPairs.Add ColPair
    
    Set ColPair = ColumnPair.Create(LHS.ListColumns.Item(3), RHS.ListColumns.Item(4))
    colPairs.Add ColPair
    
    Set ColPair = ColumnPair.Create(LHS.ListColumns.Item(4), RHS.ListColumns.Item(3))
    colPairs.Add ColPair
    
    'Assert:
    Dim Result As Variant
    Set Result = colPairs.GetPair(RHS:=RHS.ListColumns.Item(2))
    If Result Is Nothing Then GoTo TestFail
    
    Set ColPair = ColumnPair.Create(LHS.ListColumns.Item(1), RHS.ListColumns.Item(2))
    colPairs.Add ColPair
    colPairs.AddOrReplace ColPair

    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
