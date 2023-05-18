VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "City"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Example2.Model"
Option Explicit

Private Type TState
    Name As String
    Country As Country
    Airport As String
End Type
Private This As TState

'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
End Property

Public Property Get Country() As Country
    Set Country = This.Country
End Property

Public Property Set Country(ByVal vNewValue As Country)
    Set This.Country = vNewValue
End Property

Public Property Get Airport() As String
    Airport = This.Airport
End Property

Public Property Let Airport(ByVal vNewValue As String)
    This.Airport = vNewValue
End Property
