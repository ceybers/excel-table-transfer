# XMLSettingsModel
## Description
- A model to store persistent data in a Workbook, in the form of: settings, flags, and collections.
  - Settings are key value pairs, with the values always being strings. Default value is `vbNullString`.
  - Flags are boolean keys that can be either `True` or `False`. Default value is `False`.
  - Collections are collections of Strings. Default is an empty `Collection` object.
- All three of these can be stored either at the workbook-level (singleton), or at the table-level (supporting multiple tables).
- The tables are not linked or limited to the actual ListObjects in the workbook.
- If a key does not exist, the getter returns the default value. The setter will automatically insert the key if it doesn't exist, and will update it if it does (i.e., Upsert).
- If no settings model already exists, using the Create method will create an empty one.

## Sample Code
```vb
Dim s As ISettingsModel
Set s = XMLSettingsModel.Create(ActiveWorkbook, "TestSettings")
s.Workbook.GetFlag ("Foobar")

Dim SomeVariable as Boolean
SomeVariable = s.Table("Table1").GetFlag("Barfoo")
s.Table("Table1").SetFlag("Barfoo", TRUE)
```

## XMLSettingsModel methods
- `Create(ByVal Workbook As Workbook, ByVal RootNode As String) As XMLSettingsModel`
- `Reset`
- `Delete`

## ISettingsModel methods
- `Workbook() As ISettings`
- `Table(ByVal TableName As String) As ISettings`

## ISettings methods
- `GetFlag(ByVal FlagName As String) As Boolean`
- `SetFlag(ByVal FlagName As String, ByVal Value As Boolean)`
- `GetSetting(ByVal SettingName As String) As String`
- `SetSetting(ByVal SettingName As String, ByVal Value As String)`
- `GetCollection(ByVal CollectionName As String) As Collection`
- `SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)`