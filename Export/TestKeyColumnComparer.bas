Attribute VB_Name = "TestKeyColumnComparer"
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
    Dim Comparer As KeyColumnComparer
    Set Comparer = KeyColumnComparer.Create(GetLHS, GetRHS)
    
    'Comparer.LHS.PrintKeys
    'DebugPrint Comparer
    
    Dim MapResult As Variant
    MapResult = Comparer.Map
    'SubPasteMap MapResult
    
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

Private Function GetLHS() As KeyColumn
    Set GetLHS = KeyColumn.FromRange(ThisWorkbook.Worksheets.Item(2).Range("A2:A5,A14"), False)
End Function

Private Function GetRHS() As KeyColumn
    Set GetRHS = KeyColumn.FromRange(ThisWorkbook.Worksheets.Item(2).Range("C2:C13"))
End Function

Private Sub SubPasteMap(ByVal Map As Variant)
    Dim rng As Range
    Set rng = ThisWorkbook.Worksheets.Item(2).ListObjects.Item(2).ListColumns.Item(2).DataBodyRange
    Dim arr As Variant
    arr = rng.Value2
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        arr(i, 1) = Map(i + 1)
    Next i
    rng.Value2 = arr
End Sub

Private Sub DebugPrint(ByVal Compare As KeyColumnComparer)
    Debug.Print "TEST KeyColumnComparer"
    Debug.Print "===="
    Debug.Print "IsSubsetLHS = " & Compare.IsSubsetLHS
    Debug.Print "IsSubsetRHS = " & Compare.IsSubsetRHS
    Debug.Print "IsMatch = " & Compare.IsMatch
    Debug.Print "LHSOnly = " & Compare.LeftOnly.Count
    Debug.Print "RHSOnly = " & Compare.RightOnly.Count
    Debug.Print "Intersection = " & Compare.Intersection.Count
    Debug.Print vbNullString
End Sub
