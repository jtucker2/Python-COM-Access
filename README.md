# Kraken

Kraken is a tool to export elements of a Microsoft Access project in plain text. The exports can then be reimported.

# Running the tool

## Pre-requisits
- Windows
- Microsoft Access
- Python 3.10
- sqlite-tools-win32-x86

## Installing SQLite
1. Go to https://www.sqlite.org/download.html and donwload the sqlite tools for Windows.
1. Extract the files into a new folder
1. Search for environment variables in Windows and open "Edit environment variables..."
1. Click on environment variables at the bottom
1. Click on "Path" under "System variables" and click "Edit"
1. Click "New" and enter the path of where you extracted the sqlite files to

## Dumping
1. Clone the repo
1. (optional) activate the virtual environment
	```
	.\env\Scripts\activate.bat
	```
1. Install required libraries
	```
	pip install -r requirements.txt
	```
1. Get the domain model editor Access project (`DomainModeller - vX-X-X - Empty.accdb`) from sharepoint
1. Dump the access project
	
	```
	python kraken.py dump-all <access_project_file_path>
	```
	Keep access in view and press enter any time a pop-up window appears

	If no export path is specified, one will be created in Kraken directory

	The export folder should be populated with .sql, .frm and .bas files
1. Get the CSVs for the domain model by cloning https://github.com/Spyderisk/domain-network
1. Load csv data
	```
	python csv_loader.py <csv_path>
	```
1. Dump the generated sqlite database (DomainModel.db)
	1. Open the database
		```
		sqlite3 DomainModel.db
		```
	1. Set the output file
		```
		.output data.sql
		```
	1. Dump the database
		```
		.dump
		```
	1. Exit the databse
		```
  		.exit
  		```

## Loading
1. Create an empty Access project
1. Run the load command and give the directory of the empty access project
	
	```
	python kraken.py load-all <access_project_file_path>
	```
	If no export path is specified, it will be assumed there is an export folder in the Kraken directory

# All commands
```
python kraken.py [-h] [-export_path EXPORT_PATH] [-element_name ELEMENT_NAME] command project_path

command options:
	dump-all
	load-all

	dump-form -element_name ELEMENT_NAME
	load-form -element_name ELEMENT_NAME
	dump-module -element_name ELEMENT_NAME
	dump-query -element_name ELEMENT_NAME
	dump-table -element_name ELEMENT_NAME

	dump-forms
	load-forms
	dump-modules
	dump-queries
	dump-tables

	load-tables
	load-queries
```

# Info
If you get a `has no attribute 'CLSIDToClassMap'` error then delete the folder at `C:\Users\<my username>\AppData\Local\Temp\gen_py` (https://stackoverflow.com/questions/33267002/why-am-i-suddenly-getting-a-no-attribute-clsidtopackagemap-error-with-win32com)

The contents of an MS Access database can be queried using the following SQL statement:

```
SELECT MsysObjects.Name AS [List Of Tables]
FROM MsysObjects
WHERE (((MsysObjects.Name Not Like "~*") And (MsysObjects.Name Not Like "MSys*")) 
	AND (MsysObjects.Type=1))
ORDER BY MsysObjects.Name;
```

To change the type of data that is returned change the `(MsysObjects.Type)=1` part

| Type | Number |
| ---- | ------ |
| Tables | (Local):	1 |
| Tables | (Linked using ODBC):	4 |
| Tables | (Linked): 6 |
| Queries | 5 |
| Forms | -32768 |
| Reports | -32764 |
| Macros | -32766 |
| Modules | -32761 |
