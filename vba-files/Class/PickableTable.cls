VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PickableTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.TableMapper.Model"
Option Explicit

Private Type TPickableTable
    ListObject As ListObject
    Key As String
    Name As String
    Size As String
    
    ' TODO Consider impl IListable?
    Selected As Boolean
    Checked As Boolean
    Enabled As Boolean
End Type
Private This As TPickableTable

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

'@Description "a"
Public Property Get Key() As String
Attribute Key.VB_Description = "a"
    Key = This.Key
End Property

Public Property Let Key(ByVal vNewValue As String)
    This.Key = vNewValue
End Property

'@Description "Name of the underlying ListObject."
Public Property Get Name() As String
Attribute Name.VB_Description = "Name of the underlying ListObject."
    Name = This.Name
End Property

'@Description "Size of the table as a string formatted as R1C1"
Public Property Get Size() As String
Attribute Size.VB_Description = "Size of the table as a string formatted as R1C1"
    Size = This.Size
End Property

'@Description "Name of the underlying Worksheet."
Public Property Get WorksheetName() As String
Attribute WorksheetName.VB_Description = "Name of the underlying Worksheet."
    WorksheetName = This.ListObject.Parent.Name
End Property

'@Description "Name of the underlying Workbook."
Public Property Get WorkbookName() As String
Attribute WorkbookName.VB_Description = "Name of the underlying Workbook."
    WorkbookName = This.ListObject.Parent.Parent.Name
End Property

Public Property Get Selected() As Boolean
    Selected = This.Selected
End Property

Public Property Let Selected(ByVal vNewValue As Boolean)
    This.Selected = vNewValue
End Property

Public Property Get Checked() As Boolean
    Checked = This.Checked
End Property

Public Property Let Checked(ByVal vNewValue As Boolean)
    This.Checked = vNewValue
End Property

Public Property Get Enabled() As Boolean
    Enabled = This.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    This.Enabled = vNewValue
End Property

' ---
Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    This.Name = ListObject.Name
    This.Size = "R" & ListObject.DataBodyRange.Rows.Count & "C" & ListObject.DataBodyRange.Columns.Count
    This.Enabled = True
End Sub

Public Function Create(ByVal ListObject As ListObject) As PickableTable
    Dim Result As PickableTable
    Set Result = New PickableTable
    With Result
        .Load ListObject
        .Key = .WorkbookName & "::" & .WorksheetName & "::" & .Name
        If .Name = "Table2" Then .Enabled = False
    End With
    Set Create = Result
End Function
