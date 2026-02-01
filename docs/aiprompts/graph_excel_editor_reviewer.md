

# graph_excel_handler.py - Edit excels using python responsibly

## Status
[ ] Not Started | [ ] In Progress | [x] Implemented

## Summary
Under sharepoint, getting openpyxl to edit an excel when someone is coauthoring it corrupts the excel file.

It is understood that graph can solve this problem as it authenticates as an office 365 user and edits the file in a session as this, thus co authoring the file properly and causing no corruption of the file

## Context
- the Microsoft Graph API must be used 
- Excel files that are being coauthored are not corruputed by edits done by scripts

## Requirements
1. Script must be able to open an excel, read, write, trigger a recalculating of all formulas and refresh all data, save the excel, and run macros or automation scripts
2. All the above need to be implemented in a toolkit that a developer can use to safely manipulate an excel
3. pytests for the above need to be created
4. Add try/except blocks that capture the exception type, message, and traceback, then return them in a structured dict with keys: exception_type, exception_message, traceback

## Files to Create
- `lib/graph_excel_handler.py` - create the scripts here


## Testing
- write pytests to test each operation detailed in the requirements
    - The tests must set up an excel, and test the appropriate action performed on the excel
    - the test must then check that the action is successfully completed

