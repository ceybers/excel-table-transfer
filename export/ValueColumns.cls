VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.ValueMapper.Model")
Option Explicit

Private Type TValueColumns
    ListObject As ListObject
    KeyColumn As ListColumn
    Item As Scripting.Dictionary ' of type ValueColumn
    
    ShowStarredOnly As Boolean
    EnableNontext As Boolean ' Obsolete?
    EnableNonunique As Boolean ' Obsolete
    
    StarredColumns As Collection
End Type
Private This As TValueColumns

Private Sub Class_Initialize()
    Set This.Item = New Scripting.Dictionary
    Set This.StarredColumns = New Collection
End Sub

'@Enumerator
Public Property Get Item() As Scripting.Dictionary
Attribute Item.VB_UserMemId = -4
    Set Item = This.Item
End Property
' ---
Public Property Get ShowStarredOnly() As Boolean
    ShowStarredOnly = This.ShowStarredOnly
End Property

Public Property Let ShowStarredOnly(ByVal vNewValue As Boolean)
    This.ShowStarredOnly = vNewValue
End Property

Public Property Get EnableNontext() As Boolean
    EnableNontext = This.EnableNontext
End Property

Public Property Let EnableNontext(ByVal vNewValue As Boolean)
    This.EnableNontext = vNewValue
End Property

Public Property Get EnableNonunique() As Boolean
    EnableNonunique = This.EnableNonunique
End Property

Public Property Let EnableNonunique(ByVal vNewValue As Boolean)
    This.EnableNonunique = vNewValue
End Property

' ---
Public Sub Load(ByVal ListColumn As ListColumn)
    Debug.Assert Not Me Is ValueColumns
    Set This.ListObject = ListColumn.Parent
    Set This.KeyColumn = ListColumn
    
    Reload
End Sub

Public Sub Reload()
    This.Item.RemoveAll
    
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        TryAdd ValueColumn.Create(ListColumn)
    Next ListColumn
End Sub

Private Sub TryAdd(ByVal ValueColumn As ValueColumn)
    If ValueColumn.Key = This.KeyColumn.Name Then
        ValueColumn.IsKey = True
        ValueColumn.Enabled = False
    End If
    
    ValueColumn.IsStarred = CollectionEx.From(This.StarredColumns).Contains(ValueColumn.Key)
    
    ' Starred columns overrides non-enabled/ineligiblity
    If This.ShowStarredOnly Then ValueColumn.Enabled = ValueColumn.IsStarred
    
    This.Item.Add Key:=ValueColumn.Key, Item:=ValueColumn
End Sub

Public Function GetByKey(ByVal Key As String) As ValueColumn
    Set GetByKey = This.Item.Item(Key)
End Function

Public Sub UpdateStarredColumns(ByVal Collection As Collection)
    Set This.StarredColumns = Collection
    Reload
End Sub
