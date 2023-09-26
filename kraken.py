import sys
import os
import traceback
import win32com.client as win32
import pyodbc
import sqlite3

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + sys.argv[1] + ';')

project = win32.gencache.EnsureDispatch('Access.Application')
project.Application.OpenCurrentDatabase(sys.argv[1])

currentProject = project.Application.CurrentProject
currentData = project.Application.CurrentData

exportPath = sys.argv[2]

if os.path.exists("DomainModel.db"):
  os.remove("DomainModel.db")

def removExtension(fileName):
    return fileName.split(".")[0]

def dumpForm(formName):
    try:
        project.DoCmd.OpenForm(formName)
        project.Application.SaveAsText(2, formName, os.path.join(exportPath, formName + ".frm"))
        project.DoCmd.Close(2, formName)
    except:
        print("Form error", formName)
        traceback.print_exc()

def loadForm(formName):
    project.Application.LoadFromText(2, formName, os.path.join(exportPath, formName + ".frm"))

def dumpModule(moduleName):
    try:
        project.DoCmd.OpenModule(moduleName)
        project.Application.SaveAsText(5, moduleName, os.path.join(exportPath, moduleName + ".bas"))
        project.DoCmd.Close(5, moduleName)
    except:
        print("Module error", moduleName)
        traceback.print_exc()

def dumpQuery(queryName):
    dbName = project.DBEngine.Workspaces(0).Databases(0).Name
    try:
        queryString = project.DBEngine.Workspaces(0).OpenDatabase(dbName).QueryDefs(queryName).SQL
        path = os.path.join(exportPath, queryName + ".sql")
        f = open(path, "w")
        f.write(queryString)
        f.close()
    except:
        print("Query error", queryName)
        traceback.print_exc()

def dumpAllForms():
    allForms = currentProject.AllForms
    formNames = []
    for i in range(allForms.Count):
        formNames.append(allForms.Item(i).Name)

    count = 1
    for formName in formNames:
        print(str(count) + "/" + str(len(formNames)) + " forms", end = "\r")
        dumpForm(formName)
        count += 1
    print()

def loadAllForms():
    formNames = [removExtension(name) for name in os.listdir(exportPath)]

    for formName in formNames:
        print(formName)
        loadForm(formName)

def dumpAllModules():
    allModules = currentProject.AllModules
    moduleNames = []
    for i in range(allModules.Count):
        moduleNames.append(allModules.Item(i).Name)

    count = 1
    for moduleName in moduleNames:
        print(str(count) + "/" + str(len(moduleNames)) + " modules", end = "\r")
        dumpModule(moduleName)
        count += 1
    print()

def dumpAllQueries():
    allQueries = currentData.AllQueries
    queryNames = []
    for i in range(allQueries.Count):
        queryNames.append(allQueries.Item(i).Name)
    
    count = 1
    for queryName in queryNames:
        print(str(count) + "/" + str(len(queryNames)) + " queries", end = "\r")
        dumpQuery(queryName)
        count += 1
    print()

def fieldsString(fields):
    s = str(fields)
    s = s.replace("[", "(")
    s = s.replace("]", ")")
    s = s.replace("'", "")
    return s

def rowString(row):
    return str(row).replace("None", "NULL")

def decode_sketchy_utf16(raw_bytes):
    s = raw_bytes.decode("utf-16le", "ignore")
    try:
        n = s.index('\x00')
        s = s[:n]
    except ValueError:
        pass
    return s

def getFieldsAndTypes(cursor, tableName):
    fieldsList = []
    typesList = []
    for row in cursor.columns(table=tableName):
        fieldsList.append(row.column_name)
        typesList.append(row.type_name)

    s = "("
    for i in range(len(fieldsList)):
        s = s + fieldsList[i] + " " + typesList[i] + ", "
    s = s[:-2] + ")"
    return s

def dumpTable(tableName):
    conn.add_output_converter(pyodbc.SQL_WVARCHAR, decode_sketchy_utf16)

    cursor = conn.cursor()
    cursor.execute("select * from " + tableName)

    path = os.path.join(exportPath, "database-schema.sql")
    f = open(path, "a")
    f.write("CREATE TABLE " + tableName + getFieldsAndTypes(cursor, tableName) + "\n")
    f.close()
    
    con = sqlite3.connect("DomainModel.db", isolation_level=None)
    cur = con.cursor()
    cur.execute("CREATE TABLE " + tableName + getFieldsAndTypes(cursor, tableName))
    
    for row in cursor:
        row = rowString(row)
        # print("INSERT INTO " + tableName + " VALUES " + row)
        cur.execute("INSERT INTO " + tableName + " VALUES " + row)

def dumpTables():
    cursor = conn.cursor()
    # tables starting with _ not included because they werer causing errors and they don't have a csv file counterpart?
    tables = [listing[2] for listing in cursor.tables(tableType='TABLE') if listing[2].startswith("_") == False]
    for table in tables:
        dumpTable(table)

match sys.argv[3]:
    case "dump-all":
        dumpAllForms()
        dumpAllModules()
        dumpAllQueries()

    case "dump-form":
        dumpForm(sys.argv[4])
    case "load-form":
        loadForm(sys.argv[4])
    case "dump-module":
        dumpModule(sys.argv[4])
    case "dump-query":
        dumpQuery(sys.argv[4])
    case "dump-table":
        dumpTable(sys.argv[4])

    case "dump-forms":
        dumpAllForms()
    case "load-forms":
        loadAllForms()
    case "dump-modules":
        dumpAllModules()
    case "dump-queries":
        dumpAllQueries()
    case "dump-tables":
        dumpTables()

project.Application.Quit()