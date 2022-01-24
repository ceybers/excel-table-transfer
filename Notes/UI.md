# UI

## TableSelect

`() -> SelectTable -> SelectedTable`

Form should preload UI if VM properties are not null (`SelectedTable, ActiveTable, Criteria`)

## KeySelect

`(TableLHS, TableRHS, ColumnLHS) -> KeySelect -> (_, _, _, ColumnRHS)`

TableLHS as ListObject <-> ComboBox control

ListColumn in ComboBox --> Range --> ArrayExAnalyseOne --> ListView
No need to get any selection from this ListView; it is view only

ArrayIntersect --> list of strings --> ListView

 ArrayExAnalyseOne x2

 <-- User OK'ed, proceed to Value mapping
    (ShowDialog == true)
