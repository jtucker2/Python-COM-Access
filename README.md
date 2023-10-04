# Kraken

Kraken is a tool to export elements of a Microsoft Access project in plain text. The exports can then be reimported.

# Running the tool

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
1. Dump the access project
	
	(set the export path to a directory not containing kraken.py)
	```
	python kraken.py <access_project_path> <export_path> dump-all
	```
	Keep access in view and press enter any time a pop-up window appears
1. Load csv data
	```
	python kraken.py <access_project_path> <export_path> export-csvs <csv_path>
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

## Loading
1. Run the load command and give the directory of the empty access project
	
	```
	python kraken.py <access_project_path> <export_path> load-all
	```

# All commands
```
python kraken.py <project_path> <export_path>
	dump-form <form_name>
	load-form <form_name>
	dump-module <module_name>
	dump-query <query_name>
	dump-table <table_name>

	dump-forms
	load-forms
	dump-modules
	dump-queries
	dump-tables

	load-csvs <csvs_directory>
	load-tables
	load-queries

	dump-all
	load-all
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
