# VBA Automation Scripts

## Overview
This repository contains VBA scripts designed to automate various tasks in Excel, including clearing sheet contents, copying values between sheets, copying folder structures, and creating organized folders and files.

## Features
### 1. ClearSheetContents
Clears all contents below the first row for specified sheets.
- Sheets included: `Drop`, `Drop AMEX`, `Drop NAVY FCU`, `Drop Credit One`
- Displays a confirmation message after completion.

### 2. CopyValuesOnly
Copies values and formats from a source worksheet to a destination worksheet.
- Clears the destination sheet before copying.
- Preserves cell formatting.

### 3. CopyFolderStructure
Copies a folder and its contents, maintaining the structure.
- Prompts the user for source and destination paths.
- Checks if folders exist and creates them if necessary.

### 4. CreateFoldersAndFiles
Automates the creation of structured folders and Excel files.
- Prompts the user for a folder name.
- Creates subfolders for data organization.
- Copies data from predefined sheets into new Excel files.
- Saves files with specific naming conventions and optional password protection.

## Usage
1. Open the Excel file containing these macros.
2. Press `ALT + F11` to open the VBA editor.
3. Insert the code into a module if not already present.
4. Run the desired macro from the `Macros` menu (`ALT + F8`).

## Requirements
- Microsoft Excel with VBA enabled.
- Necessary permissions to create and modify files and folders.

## Contributing
Feel free to open issues or submit pull requests if you have improvements or bug fixes.

## License
This project is licensed under the MIT License.

