To run kraken.py see below:
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

Process for dumping and loading:
1. run `dump-all`
1. run `load-csvs` and give the directory of the csvs
1. dump the sqlite database into a file called `data.sql` in the same directory
1. run `load-all` and give the directory of the empty access database

Kraken may get stuck during a dump if Access is asking for input. Simply open Access and provide an input and kraken will continue.

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
