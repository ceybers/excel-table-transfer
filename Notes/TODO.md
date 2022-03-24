# TODO

## General
- [ ] Highlighting of additions/removals
- [ ] Consider adding warning for data-loss for: remove orphans, and clear destination column when not a 1:1 match
- [ ] Map values into newly created columns
- [ ] Finish mapping without any mapped columns (i.e. to add/remove keys only)
- [ ] "Transfer" sort order (i.e. Sort the keys in destination table by the order they are in the source table)
- [ ] Transfer .NumberFormat from first cell in src column to entire dst column
- [ ] Transfer ColumnWidths
- [ ] Transfer font-family, font-size, vertical-alignment, horizontal-alignment, text-wrapping
## Table Select
*None*
## Key Mapper
- [ ] Try and auto-match column when changing tables. Need to handle ComboBox changing from user, and changing from parent ComboBox resetting it.
## Column Mapper
- [ ] Highlight mapped cells, differentiating between: 0 → a, a → b, a == b, a → 0 
## Transfer History
- [ ] Transfer History fails on unsaved workbooks
- [ ] Using Split(curStr, "") in parsing history fails when filename has spaces in it
## Options
- [ ] Default Save Instruction to Yes if workbook already has a history
- [ ] Option to ignore columns with column-wide formulae in Source and/or Destination from tranfer
- [ ] Option to exclude non-unique keys from mapping (i.e. do not map to key vs map to first occurrence)
- [ ] Option to include/exclude values transfered by VarType()
## Cancelled
* ~~Map Value Columns implement multi-select on ListView without breaking single selection buttons e.g. Map~~ ListView control doesn't seem to allow check/uncheck of multiple selected list items.
## Complete
- [x] KeyColumn
- [x] ColumnPair and ColumnPairs
- [x] ValueMapper2 View and ViewModel
- [x] Options for ValueMapper
- [x] TransferTool2
- [x] Paste variant array into filtered range 
- [x] "Show mapped only" x2 for Value Mapper
- [x] Remove unmapped in source
- [x] Append unmapped in destination
- [x] Serialize TransferInstruction to/from VeryHidden worksheet
- [x] Update git README.md with new screenshots
- [x] Add splash screen that asks if the table in Selection is the source or destination. This replaces TableSelect as the intial dialog
- [x] Match key columns by name
- [x] Auto-select table if only one other table available to select, and focus on OK/Next
- [x] Check if OK button works like double clicking treeview
- [x] Change icons in Key Column Mapper to reflect test results, i.e. green check only if Unique = Count
- [x] Rename Key Set Theory ListView headers to Orphans and Additions
- [x] Suppress prompt to auto-match key column by name when first showing dialog (and before user input)
- [x] Cancel/Back/Next/Finish stages (AppContext?)
- [x] Fail gracefully if only one table available
- [x] Add column number (i.e. Column A) to Column Mapping dialog
- [x] Default Respect Filters to Yes for both Source and Destination
- [x] Highlighting of mapped value cells
- [x] Flag columns with column-wide formulae in ListView
- [x] Add options for insert/removing keys under Key Set Theory
- [x] Smart key column guessing (try to find columns with 100% unique values and 1:1 match between src/dst)