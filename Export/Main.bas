Attribute VB_Name = "Main"
'@Folder "TableTransferTool"
Option Explicit

'@ExcelHotkey e
'@EntryPoint
Public Sub RunTableTransferTool()
Attribute RunTableTransferTool.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim AppContext As AppContext
    Set AppContext = New AppContext
    AppContext.Start
End Sub
