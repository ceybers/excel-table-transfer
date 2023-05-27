Attribute VB_Name = "CollectionExConverters"
'@Folder("Helpers.CollectionEx")
Option Explicit

Private Const DEFAULT_DELIMITER As String = ","

Public Sub TestCollectionConverters()
    Dim d As Dictionary
    Set d = TextToDictionary("a,b,c")
    Stop
End Sub

Public Function TextToDictionary(ByVal Text As String, Optional ByVal Delimiter As String) As Scripting.Dictionary
    Debug.Assert Text <> vbNullString
    Debug.Assert Len(Delimiter) < 2
    
    Dim SplitText As Variant
    SplitText = Split(Text, IIf(Delimiter = vbNullString, DEFAULT_DELIMITER, Delimiter))
    
    Debug.Assert UBound(SplitText) > 0
    
    Set TextToDictionary = ArrayEx.From(SplitText).ToDictionary
End Function
