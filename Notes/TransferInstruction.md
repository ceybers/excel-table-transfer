# Transfer Instruction

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