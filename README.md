# Excel Maximum Value Extractor_v2.0
A Python application to extract the maximum values from specified columns in multiple Excel files based on user-defined conditions.

## Table of Contents
- Introduction
- Features
- Requirements
- Installation
- Usage
- Example
- Contributing
- License
- Contact Information
- Acknowledgments
- Change Log

## Introduction
The Excel Maximum Value Extractor is a GUI-based Python application that allows users to process multiple Excel files to extract maximum values from a specified target column, only when a condition in another column is met. The application supports batch processing, user input for column names, and provides progress updates during the processing.

## Features
- User-Defined Condition and Target Columns: Input custom column names for condition and target columns.
- Data Filtering: Extract maximum values from rows where the condition column values exceed a specified threshold (default is >10).
- Batch Processing: Select and process multiple Excel files in one go.
- Progress Window: Displays the current file and worksheet being processed, along with a progress bar and estimated remaining time.
- Error Handling: Handles missing columns, empty data, and other exceptions gracefully, providing informative messages.
- Customizable Output: Saves the results with default filenames based on the original Excel files, which can be customized by the user.
- User Interaction: Provides options to retry or exit when encountering issues like missing input or canceled operations.
- Platform Compatibility: Uses standard Python libraries for GUI (tkinter) and data handling (pandas), ensuring cross-platform compatibility.

## Requirements
- Python 3.x: The application requires Python 3.x to run.
- pandas Library: For data manipulation and Excel file processing.
- tkinter Library: For GUI components (usually included with Python installations).

## Installation
1. Clone the Repository (if applicable):

```bash
git clone https://github.com/yourusername/yourrepository.git
```
2. Install Required Libraries:

   If you don't have pandas installed, install it using:

```bash
pip install pandas
```

The tkinter library is typically included with Python, but if not, install it accordingly for your operating system.


## Usage
1. Run the Application:
   - You can run the script directly if it's a .pyw file (which hides the console window) by double-clicking it.
   - Alternatively, run it from the command line:

```bash
python your_script_name.py
```
2. Input Condition and Target Columns:

   - Condition Column: When prompted, enter the name of the column to use as the condition for filtering (default is "Ch2, x").
   - Target Column: Next, enter the name of the column from which to extract the maximum value (default is "Ch2, y").

3. Select Excel Files:
   - A file dialog will appear. Select one or more Excel files (.xlsx format) you wish to process.

4. Processing:
    - The application will process each selected file.
    - A progress window will display:
        - Current File: The name of the file being processed.
        - Current Worksheet: The name of the worksheet being processed.
        - Progress Bar: Visual indicator of processing progress.
        - Estimated Remaining Time: Approximate time left to complete processing.

5. Save Results:
    - After processing each file, you will be prompted to save the results.
    - A default filename is suggested (original filename with _結果.xlsx appended).
    - Choose the save location and filename, or accept the default.

6. Completion:
    - After all files are processed, a message will inform you that processing is complete.
    - The application will then exit.

## Example
Suppose you have multiple Excel files with worksheets containing data columns "Ch2, x" and "Ch2, y". You want to extract the maximum values from "Ch2, y" where "Ch2, x" values are greater than 10.

1. Input:
   - Condition Column: Ch2, x
   - Target Column: Ch2, y

2. Process:
The application filters each worksheet for rows where Ch2, x > 10.
It then calculates the maximum value of Ch2, y from the filtered data.

3. Output:
   - A new Excel file for each input file containing:
   - A sheet listing each worksheet's name and the corresponding maximum value from Ch2, y.

## Contributing
Contributions are welcome! Please follow these steps:

1. Fork the Repository: Create your own fork of the project.
2. Create a Branch: Make a new branch for your feature or bugfix.
```bash
git checkout -b feature/your-feature-name
```

3. Commit Your Changes: Write clear and concise commit messages.

4. Push to Your Fork:
```bash
git push origin feature/your-feature-name
```

5. Submit a Pull Request: Describe your changes and submit the PR for review.

## License
MIT License

## Contact Information
- Author: Elite Lu
- GitHub: https://github.com/imelitelu

## Acknowledgments
- pandas: For data manipulation and Excel file handling.
- tkinter: For the GUI components.
- Community: Thanks to all users and contributors for their feedback and improvements.

## Change Log

### v1.0 feat: add custom column input, data filter, improved UX

 Add custom input for condition and target columns, implement data filtering, improve error handling and user experience

- Allow users to input condition and target column names
- Filter data based on condition column values (>10)
- Display current sheet name in the progress window
- Enhance error handling for missing columns and other exceptions
- Provide the ability to exit during processing
- Improve user interaction with additional prompts


### v1.1 fix: unified exit prompts for column input dialogs
Unified Exit Handling for Column Input with Retry Option
- Updated exit prompt for condition and target column inputs to use askretrycancel, allowing users to retry input instead of exiting immediately.
- Improved consistency in user prompts and exit handling across the program.
- Updated file naming method to remove version suffix from the current active file.
- Added "Archive" folder to store old versions.

### v2.0 feat: add multi-file selection and file name display in progress
Enhanced Progress Window for Clarity and Multi-File Tracking

- Display Current File Name in Progress Window: Added a label above the current sheet name in the progress window to show the file currently being processed, giving users better visibility.
- Optimized Progress Window Layout: Adjusted the order of elements in the progress window to display the file name, current sheet name, and progress bar in sequence, improving user experience.
- Enhanced Multi-File Processing: The progress window now shows the current file and sheet being processed for each file, making it easier for users to follow progress.
- Change module from "Vibration_analysis" to "Excel Maximum Value Extractor"