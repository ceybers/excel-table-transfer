VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.KeyMapper.Model"
Option Explicit

Private Type TKeyColumns
    ListObject As ListObject
    Item As Scripting.Dictionary ' of type PickableTable
    EnableNontext As Boolean
    EnableNonunique As Boolean
End Type
Private This As TKeyColumns

Private Sub Class_Initialize()
    Set This.Item = New Scripting.Dictionary
End Sub

'@Enumerator
Public Property Get Item() As Scripting.Dictionary
Attribute Item.VB_UserMemId = -4
    Set Item = This.Item
End Property
' ---
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
Public Sub Load(ByVal ListObject As ListObject)
    Debug.Assert Not Me Is KeyColumns
    Set This.ListObject = ListObject
    
    Reload
End Sub

Public Sub Reload()
    This.Item.RemoveAll
    
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        TryAdd KeyColumn.Create(ListColumn)
    Next ListColumn
End Sub

Private Sub TryAdd(ByVal KeyColumn As KeyColumn)
    KeyColumn.Enabled = True
    
    If (Not This.EnableNontext) And (Not KeyColumn.IsText) Then KeyColumn.Enabled = False
    If (Not This.EnableNonunique) And (Not KeyColumn.IsUnique) Then KeyColumn.Enabled = False
    
    This.Item.Add Key:=KeyColumn.Key, Item:=KeyColumn
End Sub
