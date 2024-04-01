Attribute VB_Name = "StringConstants"
'@Folder "MVVM.Resources.Constants"
Option Explicit

Public Const APP_TITLE As String = "Table Transfer Tool"
Public Const APP_VERSION As String = "Version 1.8.4-dev"
Public Const APP_COPYRIGHT As String = "2024 Craig Eybers" & vbCrLf & "All rights reserved."

Public Const TAG_WORKBOOK As String = "WORKBOOK"
Public Const TAG_TABLE As String = "TABLE"

Public Const LAST_KEY_USED As String = "LastKeyUsed"

Public Const MSG_CAPTION As String = "Table Transfer Tool"

Public Const NO_TABLES_FOUND As String = "(No tables open in Excel)"
Public Const NO_COLUMN_SELECTED As String = "(No column selected)"
Public Const NO_TWO_COLUMNS_SELECTED As String = vbNullString
Public Const NO_TABLE_SELECTED As String = "(No table selected)"

Public Const MSG_ZERO_KEYS As String = "Zero keys found!"

Public Const KEY_HEADER As String = "Keys"
Public Const FIELD_HEADER As String = "Fields"
Public Const SELECT_ALL As String = "(Select all)"

Public Const ERR_CAPTION As String = "#ERROR"
Public Const ERR_SOURCE As String = "Table Transfer Tool"
Public Const ERR_MSG_GENERIC As String = "An unexpected error occurred." & vbCrLf & "Table Transfer Tool will now close."

'Public Const NUM_FMT_NUMBER As String = "Standard" '"0.00"
'Public Const NUM_FMT_CURRENCY As String = "Currency" '"$# ##0.00"
'Public Const NUM_FMT_DATE As String = "Long Date" '"yyyy/mm/dd"

Public Const HDR_TXT_VALUE_MAPPER As String = "Which columns contain the data you want to transfer?" & vbCrLf & vbCrLf & _
    "Select pairs of columns from the Source and Destination tables and add them to the Mapped columns."

Public Const HDR_TXT_TABLE_PICKER As String = "Which two tables are you transferring data between?" & vbCrLf & vbCrLf & _
    "Select a Source table to copy data from, and a Destination table to insert and update data."

Public Const HDR_TXT_KEY_MAPPER As String = "Which two columns contain the primary key used to join your tables?" & vbCrLf & vbCrLf & _
    "Select a Key column in both the Source and the Destination table."

Public Const HDR_TXT_DELTAS_PREVIEW As String = "Table differences compared successfully." & vbCrLf & vbCrLf & _
    "The changes can be previewed below before applying them to the Destination table."

Public Const DELTAS_PREVIEW_NO_RESULTS As String = "Table differences compared successfully." & vbCrLf & vbCrLf & _
    "No changes were found."


Public Const CHANGE_TYPE_INVALID As String = "CHANGE_TYPE_INVALID"
Public Const CHANGE_TYPE_BLANK_UNCHANGED As String = "CHANGE_TYPE_BLANK_UNCHANGED"
Public Const CHANGE_TYPE_VALUE_REPLACES_BLANK As String = "CHANGE_TYPE_VALUE_REPLACES_BLANK"
Public Const CHANGE_TYPE_BLANK_REPLACES_VALUE As String = "CHANGE_TYPE_BLANK_REPLACES_VALUE"
Public Const CHANGE_TYPE_VALUE_UNCHANGED As String = "CHANGE_TYPE_VALUE_UNCHANGED"
Public Const CHANGE_TYPE_VALUE_CHANGED As String = "CHANGE_TYPE_VALUE_CHANGED"

Public Const MAGIC_FORMULA_HIGHLIGHTING As String = "=OR(TRUE,""HighlightMapped;b92d7b59-e7ec-4db0-a7c6-5a6ad86ceac2"")"

Public Const COLOR_GREEN_DARK As Long = 14348258 'RGB(226, 239, 218)
Public Const COLOR_GREEN_LIGHT As Long = 10092492 'RGB(204, 255, 153) '#CCFF99
Public Const COLOR_DEFAULT_HIGHLIGHT As Long = 10092492 'RGB(204, 255, 153) '#CCFF99
Public Const COLOR_NO_COLUMN_SELECTED As Long = 8421504 'RGB(128,128,128)
Public Const COLOR_NO_TABLES_AVAILABLE As Long = 8421504 ' RGB(128, 128, 128)
Public Const COLOR_SELECT_ALL As String = 8421504 'RGB(128, 128, 128)

Public Const ERR_NUM_KEYCOL_EMPTY_TABLE As Long = vbObjectError + 1
Public Const ERR_MSG_KEYCOL_EMPTY_TABLE As String = "ERR_KEYCOL_EMPTY_TABLE"
Public Const ERR_NUM_KEYCOLCOMP_EMPTY_TABLE As Long = vbObjectError + 2
Public Const ERR_MSG_KEYCOLCOMP_EMPTY_TABLE As String = "ERR_KEYCOLCOMP_EMPTY_TABLE"
Public Const ERR_NUM_MULTIPLE_COLUMNS As Long = vbObjectError + 3
Public Const ERR_MSG_MULTIPLE_COLUMNS As String = "Cannot create clsKeyColumn with a range that spans multiple columns"
Public Const ERR_NUM_NO_VISIBLE_CELLS As Long = vbObjectError + 4
Public Const ERR_MSG_NO_VISIBLE_CELLS As String = "KeyColumn.LoadRange failed"
Public Const ERR_NUM_UNEXPECTED_VARTYPE As Long = vbObjectError + 5
Public Const ERR_MSG_UNEXPECTED_VARTYPE As String = "Unexpected VarType TransferDeltas"
