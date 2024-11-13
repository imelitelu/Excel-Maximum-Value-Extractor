# Introduction
This is a analysis module for pick up the maxmun value in which you decide to pick up row.

---

# Change Log
## V1 feat: add custom column input, data filter, improved UX

 Add custom input for condition and target columns, implement data filtering, improve error handling and user experience

- Allow users to input condition and target column names
- Filter data based on condition column values (>10)
- Display current sheet name in the progress window
- Enhance error handling for missing columns and other exceptions
- Provide the ability to exit during processing
- Improve user interaction with additional prompts


## V2 fix: unified exit prompts for column input dialogs
Unified Exit Handling for Column Input with Retry Option
- Updated exit prompt for condition and target column inputs to use askretrycancel, allowing users to retry input instead of exiting immediately.
- Improved consistency in user prompts and exit handling across the program.
- Updated file naming method to remove version suffix from the current active file.
- Added "Archive" folder to store old versions.

## V3 feat: add multi-file selection and file name display in progress
Enhanced Progress Window for Clarity and Multi-File Tracking

- Display Current File Name in Progress Window: Added a label above the current sheet name in the progress window to show the file currently being processed, giving users better visibility.
- Optimized Progress Window Layout: Adjusted the order of elements in the progress window to display the file name, current sheet name, and progress bar in sequence, improving user experience.
- Enhanced Multi-File Processing: The progress window now shows the current file and sheet being processed for each file, making it easier for users to follow progress.