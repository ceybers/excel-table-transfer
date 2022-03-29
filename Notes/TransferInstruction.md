# Transfer Instruction

## Naming Restrictions

### Tables
* Must start with A-z, _ or \
* Rest is A-z, 0-9, ., _
* Cannot use names C, c, R, r
* Cannot be the same as cell references
* Cannot use spaces
* Maximum length 255

### Columns
* Special characters: tab, LF, CR, ,:.[]#'"{}$^&*+==<>/
* Require escaping: [, ], #, '
## ?
Source ListObject
Destination ListObject
SourceKey ListColumn
DestinationKey ListColumn
MapPairs Collection<T>
   Src ListColumn
   Dst ListColumn
Options
    SourcingStrategy
        FilteredOnly
        IgnoreBlanks
    TransferStrategy
        FilteredOnly
        ClearDestinationFirst
        ReplaceEmptyOnly
    TransferType
        AsValue
        AsFormula
        AsR1C1
    Keys
        IgnoreAdditions
        IgnoreRemovals
        MapMatches
    Formatting
        HighlightAdditionKeyOnly
        HighlightAdditionKeyAndMapped
        HighlightAdditionEntireRow
        ...RemovalKeyOnly
        ...RemovalKeyAndMapped
        ...RemovalEntireRow
        etc.