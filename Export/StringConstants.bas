Attribute VB_Name = "StringConstants"
'@Folder "MVVM.Model.Constants"
Option Explicit

Public Const TAG_WORKBOOK As String = "WORKBOOK"
Public Const TAG_TABLE As String = "TABLE"

Public Const NOT_FOUND_COLOR As Long = 8421504 ' RGB(128, 128, 128)
Public Const NO_TABLES_FOUND As String = "(No tables found)"

Public Const NO_COLUMN_SELECTED As String = "(No column selected)"
Public Const NO_COLUMN_COLOR As Long = 8421504 'RGB(128,128,128)

Public Const NO_TWO_COLUMNS_SELECTED As String = vbNullString
Public Const MSG_ZERO_KEYS As String = "Zero keys found!"

Public Const NO_TABLE_SELECTED As String = "(No table selected)"

Public Const KEY_HEADER As String = "Keys"
Public Const FIELD_HEADER As String = "Fields"
Public Const SELECT_ALL As String = "(Select all)"
Public Const SELECT_ALL_COLOR As String = 8421504 'RGB(128, 128, 128)

Public Const ERR_CAPTION As String = "#ERROR"

Public Const NUM_FMT_NUMBER As String = "Standard" '"0.00"
Public Const NUM_FMT_CURRENCY As String = "Currency" '"$# ##0.00"
Public Const NUM_FMT_DATE As String = "Long Date" '"yyyy/mm/dd"

Public Const HDR_TXT_VALUE_MAPPER As String = "Which columns contain the data you want to transfer?" & vbCrLf & vbCrLf & _
    "Select pairs of columns from the Source and Destination tables and add them to the Mapped columns."

Public Const HDR_TXT_TABLE_PICKER As String = "Which two tables are you transferring data between?" & vbCrLf & vbCrLf & _
    "Select a Source table to copy data from, and a Destination table to insert and update data."

Public Const HDR_TXT_KEY_MAPPER As String = "Which two columns contain the primary key used to join your tables?" & vbCrLf & vbCrLf & _
    "Select a Key column in both the Source and the Destination table."

Public Const HDR_TXT_DELTAS_PREVIEW As String = "Table differences compared successfully." & vbCrLf & vbCrLf & "The changes can be previewed below before applying them to the Destination table."
Public Const DELTAS_PREVIEW_NO_RESULTS As String = "Table differences compared successfully." & vbCrLf & vbCrLf & "No changes were found."

Public Const ERR_SOURCE As String = "Table Transfer Tool"


