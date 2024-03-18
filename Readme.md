# Notes Import Processing Script

This repository contains a script for processing notes from an older Excel file format (.xls) and preparing them for re-import into a system. The script reads the data, applies necessary transformations, and outputs a CSV file ready for import.

## Features

- **Remapping Functionality**: Update note IDs based on a predefined mapping.
- **Data Concatenation**: Merge relevant columns into a single notes field.
- **File Creation**: Generate individual parish files and a summary of the total count.
- **Cleanup**: Remove specified columns and handle empty files or those without descriptions.
- **Import File Preparation**: Create a separate file for re-import with specified modifications.

## Usage

1. Ensure you have Python installed on your system.
2. Install required Python packages using `pip`:

   ```sh
   pip install pandas xlrd


3. Place the script in your project directory.
4. Ensure that your input file, `NotesAllChanges.XLS`, is in the same directory as the script or update the script with the correct file path.
5. Run the script:
   ```sh
   python noteSync_changeLog.py

## Functions

- `remap_fund_id()`: Remap fund IDs based on custom logic.
- `apply_remap()`: Apply remapping across DataFrame.
- `concatenate_columns()`: Concatenate specified columns for notes.
- `creating_Import_files()`: Prepare data for re-import by setting a title, dropping columns, and saving as CSV.
- `parish_files()`: Create individual parish files based on note descriptions.
- `file_clean_up()`: Clean up files in the directory, merge similar files, and create a total count summary.

## Output Files

- `Updated_NotesAllChanges.xlsx`: Intermediary file with concatenated columns.
- `NotesAllChanges_ImportReady.csv`: Final output ready for re-import, with unnecessary columns removed and 'CnNote_1_01_Title' set to 'Synced'.
- `TotalCount.xlsx`: Summary file with the count of rows for each parish file.
- `Rows Without File.xlsx`: Contains any rows that did not have a 'CnNote_1_01_Description' and could not be parsed into their own file.
- Individual parish files named after descriptions within 'CnNote_1_01_Description'.

## Important Notes

- Before running the script, make sure to fill in the `remap_dict` with your specific mapping values.
- The column names used for concatenation and removal are based on the provided script and may need to be adjusted according to your actual Excel file structure.

## Contributing

Feel free to fork this repository and submit pull requests for any improvements or fixes you make to the script.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
