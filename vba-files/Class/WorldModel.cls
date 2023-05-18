VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorldModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Example2.Model"
Option Explicit

Private Type TState
    Countries As Scripting.Dictionary
    Cities As Scripting.Dictionary
End Type
Private This As TState

Public Property Get Countries() As Scripting.Dictionary
    Set Countries = This.Countries
End Property

Public Property Get Cities() As Scripting.Dictionary
    Set Cities = This.Cities
End Property

Private Sub Class_Initialize()
    Set This.Countries = New Scripting.Dictionary
    Set This.Cities = New Scripting.Dictionary
End Sub

Public Sub Load()
    LoadCities
End Sub

Private Sub LoadCities()
    CreateCity "Atlanta", "USA", "ATL/KATL"
    CreateCity "Dallas", "USA", "DFW/KDFW"
    CreateCity "Denver", "USA", "DEN/KDEN"
    CreateCity "Dubai", "UAE", "DXB/OMDB"
    CreateCity "Istanbul", "Turkey", "IST/LTFM"
    CreateCity "London", "UK", "LHR/EGLL"
    CreateCity "Delhi", "India", "DEL/VIDP"
    CreateCity "Madrid", "Spain", "MAD/LEMD"
End Sub

Private Function CreateCity(ByVal CityName As String, ByVal CountryName As String, ByVal Airport As String) As City
    Dim ThisCountry As Country
    If This.Countries.Exists(CountryName) Then
        Set ThisCountry = This.Countries.Item(CountryName)
    Else
        Set ThisCountry = CreateCountry(CountryName)
    End If
    
    Dim NewCity As City
    Set NewCity = New City
    NewCity.Name = CityName
    Set NewCity.Country = ThisCountry
    NewCity.Airport = Airport
    This.Cities.Add Key:=CityName, Item:=NewCity
    Set CreateCity = NewCity
End Function

Private Function CreateCountry(ByVal CountryName As String) As Country
    Dim NewCountry As Country
    Set NewCountry = New Country
    NewCountry.Name = CountryName
    This.Countries.Add Key:=CountryName, Item:=NewCountry
    Set CreateCountry = NewCountry
End Function
