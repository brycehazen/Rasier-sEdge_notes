# Notes Import Processing Script:

This repository contains a script for processing notes from an older Excel file format (.xls) and preparing them for re-import into a system. The script reads the data, applies necessary transformations, and outputs a CSV file ready for import. This is used in Raiser's Edge to create a log file on the record of when changes are made on a weekly basis. The Raiser's Edge database this was scripted for used ConsCodes and matching Fund IDs to identify entities. These ensure the constituent is linked via their constituent and gift record to the entity(Parish, DAF, ect.).

## Features:

- **Remapping Functionality**: Update note IDs based on a predefined mapping.
- **Data Concatenation**: Merge relevant columns into a single notes field.
- **File Creation**: Generate individual parish files and a summary of the total count.
- **Cleanup**: Remove specified columns and handle empty files or those without descriptions.
- **Import File Preparation**: Create a separate file for re-import with specified modifications.

## Set up:

* Ensure you have Python installed on your system. You will need the following libraries: 
  ```sh
      pip install pandas
      pip install xlrd
* Whenever changes are made to a record, a note with type `Sync` is added to the record, leaving all other fields blank
     * A note must only be left when all fields on a record have been screened for accuracy and meets data entry standards
* The query in folder `Database Cleanup` has the query `Notes Sync Change Log`
   * This query criteria `(Note Type =  Sync ) AND (Title, Description, Notes = Blank )`
* Export in folder `Data Integrity` called `Notes Sync Change Log`
   * This export has the following headers:

   | CnAls_1_01_Alias | CnAls_1_01_Alias_Type | LastChangedBy | DateLastChange | ConsID | IsInactive | Deceased | DeceasedDate | Gender | Titl1 | FirstName | MiddleName | LastName | Suffix | MrtlStat | MaidenName | Bday | SRConsID | SRInactive | SRDeceased |    SRDeceasedDate | SRGender | SRTitl1 | SRFirstName | SRMiddleName | SRLastName | SRSuffix | SRMrtlStat | SRMaidenName |
   | ----------------- | --------------------- | ------------- | -------------- | ------ | ---------- | -------- | ------------- | ------ | ----- | --------- | ---------- | -------- | ------ | -------- | ---------- | ---- | -------- | ---------- | ---------- | -------------- | -------- | ------- | ----------- | ------------ | ---------- | -------- | ----------- | ------------- |
   | Text              | Table                 | Sytem         | Date           | Text   | TF         | TF       | Date          | Table  | Table | Text      | Text       | Text     | Table  | Table    | Text       | Date | Text     | TF         | TF         | Date             | Table    | Table   | Text        | Text         | Table      | Table    | Text       | Text         |

   | PrimAddText | PrimSalText | AddrLines | AddrCity | AddrState | AddrZIP | CnNote_1_01_Type | CnNote_1_01_Title | CnNote_1_01_Description | CnNote_1_01_Import_ID | ConsCode_Long | ConsCode_Short | ConsCode_StartDate | ConsCode_EndDate | ConsCode_Long_1 |    ConsCode_Short_1 | ConsCode_StartDate_1 | ConsCode_EndDate_1 | ConsCode_Long_2 | ConsCode_Short_2 | ConsCode_StartDate_2 | ConsCode_EndDate_2 | LastFundIDDesc | LastFundID |
   | --------- | ----------- | --------- | -------- | --------- | -------- | ----------------- | ----------------- | ----------------------- | --------------------- | ------------- | -------------- | ------------------ | ---------------- | -------------- | ---------------- | ------------------- | ----------------- | -------------- | ---------------- | ------------------- | ----------------- | --------------- | --------- |
   | Table     | Table       | Text      | Text     | Text      | Table    | Text              | Table             | Text                    | Text                  | System        | Table          | Table              | Date             | Date           | Table            | Table               | Date              | Date            | Table             | Table                | Date                   | Table             | Text      |




## Usage

1. Place the script `noteSync_changeLog.py` in your project directory.
2. Export `Note sync change log` from the `Database Cleanup` folder, creating file `NotesAllChanges.XLS`.
3. Open `NotesAllChanges.XLS` and double check all information is correct and no data entries have been made.
   1. Make any changes as needed, this file will eventually be parsed and imported back into Raiser's Edge.
5. Ensure that your input file, `NotesAllChanges.XLS`, is in the same directory as the script.
6. Run the script:
   ```sh
   python noteSync_changeLog.py
7. This will create a folder with parsed-out files for each entity. Drop this folder into SharePoint so that each entity can make the changes in their database as well.
   1. This is done on a weekly basis.
   2. This folder has the files:
      * `Rows Without File.xlsx`: This will catch any outliers that can not be mapped to an entity
      * `TotalCount.xlsx`: logging how many rows are in each entity file
      
8. `NotesAllChanges_ImportReady.csv` is created and is ready to be imported back in
9.  Under Admin > Import > Constituent Import  - Open the import `Note syncing - record update log`
   1. Validate the import first, resolving any exceptions before importing
   2. Save an exception file
   3. Save output query in the query folder 'Note syncing - Record Update Logs'
   

## Functions

- `remap_fund_id()`: Remap fund IDs to match more closely with ConsCodes.
- `apply_remap()`: Apply remapping across DataFrame.
- `concatenate_columns()`: Concatenate specified columns for notes, which are then put back onto the record.
- `creating_Import_files()`: Prepare data for re-import by setting a title, dropping columns, and saving as CSV.
- `parish_files()`: Create individual parish files based on note descriptions.
- `file_clean_up()`: Clean up files in the directory, merge similar files, and create a total count summary.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
