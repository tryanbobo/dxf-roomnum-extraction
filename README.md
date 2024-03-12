# DXF Room Number Extractor
This Python script extracts room numbers from DXF (Drawing Exchange Format) files and organizes the data into an Excel spreadsheet for further analysis or use.

## Features
Extracts room numbers from text entities (TEXT and MTEXT) in DXF files.
Determines the floor number based on the most common leading digit in the room numbers.
Handles various room number formats and skips excluded substrings.
Generates an Excel spreadsheet with columns for "Building", "Floor", and "Room".

## Usage
Place your DXF files in a folder.

Update the root_folder variable in the script to point to your folder.

Run the script.

## Requirements
Python 3.x

ezdxf library (install with pip install ezdxf)

pandas library (install with pip install pandas)

## Example Output
The script will generate an Excel file for each DXF file processed, containing the following columns:

Building: Name of the DXF file.

Floor: Floor number based on the most common leading digit.

Room: Extracted room numbers.

## Notes:

Ensure that the DXF files contain room numbers as text entities (TEXT or MTEXT).
Modify the excluded_substrings list in extract_room_numbers to exclude specific substrings from room numbers.
## License
[MIT](https://choosealicense.com/licenses/mit/)
