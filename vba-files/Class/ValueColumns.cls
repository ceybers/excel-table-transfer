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
    
    EnableNontext As Boolean
    EnableNonunique As Boolean
End Type
Private This As TValueColumns

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
    'ValueColumn.Enabled = True
    'If (Not This.EnableNontext) And (Not ValueColumn.IsText) Then ValueColumn.Enabled = False
    'If (Not This.EnableNonunique) And (Not ValueColumn.IsUnique) Then ValueColumn.Enabled = False
    If ValueColumn.Key = This.KeyColumn.Name Then
        ValueColumn.IsKey = True
        ValueColumn.Enabled = False
    End If
    
    This.Item.Add Key:=ValueColumn.Key, Item:=ValueColumn
End Sub

Public Function GetByKey(ByVal Key As String) As ValueColumn
    Set GetByKey = This.Item.Item(Key)
End Function
