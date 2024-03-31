# Change Log
## v1.8.2-dev (2024/03/31)
- Overhauled all the UI dialogs.
- Fixed error when trying to transfer data with no matching keys.
- Fixed some weird edge case bugs.
 
## v1.8.1-dev (2024/03/30)
- Completely rewrote Transfer code. Faster, more efficient, more easily extensible.
- Started refactoring legacy UI to MVVM pattern.
- Added rudimentary Change Data Capture tracking.
- Preview changes before Committing them to the Destination.
- Recall most recent Transfer where both Tables are currently open.
- Transfer Column Width and Number Formatting to Destination.
- FIX Button logic with AutoMap Value columns feature.
- General housekeeping. Removed obsolete code, cleaned up formatting, rearranged RubberDuck folder structure and annotations.

## v1.7.0 (2023/11/27)
- Store list of user-defined keys to use with AutoGuess on Key Mapping.
- Fixed bug with persistent User Settings.
 
## v1.6.0 (2023/05/27)
- Starting completing redoing all the UI but never managed to ship it to production.

## v1.5.0 (2022/03/31)
 - Recall last Transfer on the active Workbook if the second Table is also open.
  
## v1.4.0 (202/02/25)
- Dialog box to choose if the selected Table is the Source or Destination.

## v1.3.0 (2022/01/25)
- Implemented Transfer Options: behaviour for handling filtered rows in Source, filtered Rows in Destination, Blank cells in Source, Blank cells in Destination, and whether to clear entire columns before transferring.
 
## v1.2.0 (2022/01/25)
- Lots of behind-the-scenes refactoring.

## v1.0.0 (2021/12/01)
- Initial commit 🎂