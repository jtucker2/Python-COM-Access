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

	dump-all
```

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