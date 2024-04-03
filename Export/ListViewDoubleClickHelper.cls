VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListViewDoubleClickHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Helpers.Objects")
Option Explicit

Private mMouseDoubleClicked As Boolean

Public Function TryGetDoubleClickedListItem(ByVal ListView As ListView, ByVal x As Long, ByVal y As Long, ByRef OutListItem) As Boolean
    If Not mMouseDoubleClicked Then Exit Function
    
    mMouseDoubleClicked = False
    Dim Result As ListItem
    Set Result = ListView.HitTest(x * TWIPS_PER_PIXEL, y * TWIPS_PER_PIXEL)
    If Result Is Nothing Then Exit Function
    
    Set OutListItem = Result
    TryGetDoubleClickedListItem = True
End Function

Public Sub OnDblClick()
    mMouseDoubleClicked = True
End Sub
