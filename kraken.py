import sys
import os
import traceback
import win32com.client as win32
import pyodbc
import sqlite3
import argparse


parser = argparse.ArgumentParser()
parser.add_argument('command')
parser.add_argument('project_path')
parser.add_argument(
    '-export_path', 
    default="exports", 
    help="The path for exporting/importing files."
)
parser.add_argument(
    '-element_name', 
    help="The name of an individual element to be exported."
)
args = parser.parse_args()

project = win32.gencache.EnsureDispatch('Access.Application')
project.Application.OpenCurrentDatabase(os.path.abspath(args.project_path))

currentProject = project.Application.CurrentProject
currentData = project.Application.CurrentData

exportPath = os.path.abspath(args.export_path)
try:
   os.makedirs(exportPath)
except FileExistsError:
   pass

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
        print("{}/{} forms".format(count, len(formNames)), end= "\r")
        dumpForm(formName)
        count += 1
    print()

def dumpAllModules():
    allModules = currentProject.AllModules
    moduleNames = []
    for i in range(allModules.Count):
        moduleNames.append(allModules.Item(i).Name)

    count = 1
    for moduleName in moduleNames:
        print("{}/{} modules".format(count, len(moduleNames)), end= "\r")
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
        print("{}/{} queries".format(count, len(queryNames)), end= "\r")
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

def getFieldsAndTypes(cursor, tableName):
    fieldsList = []
    typesList = []
    for row in cursor.columns(table=tableName):
        fieldsList.append(row.column_name)
        typesList.append(row.type_name)

    s = "("
    for i in range(len(fieldsList)):
        s = s + "{field} {type}, ".format(field=fieldsList[i], type=typesList[i])
    s = s[:-2] + ")"
    return s

def decode_sketchy_utf16(raw_bytes):
    s = raw_bytes.decode("utf-16le", "ignore")
    try:
        n = s.index('\x00')
        s = s[:n]
    except ValueError:
        pass
    return s

def dumpTable(tableName):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + os.path.abspath(args.project_path) + ';')

    # Converter added due to decode error - https://github.com/mkleehammer/pyodbc/issues/328#issuecomment-419655266
    conn.add_output_converter(pyodbc.SQL_WVARCHAR, decode_sketchy_utf16)

    cursor = conn.cursor()

    # path = os.path.join(exportPath, "database-schema.sql")
    # f = open(path, "a")
    # f.write("CREATE TABLE " + tableName + getFieldsAndTypes(cursor, tableName) + "\n")
    # f.close()
    
    con = sqlite3.connect("DomainModel.db", isolation_level=None)
    cur = con.cursor()
    cur.execute("CREATE TABLE " + tableName + getFieldsAndTypes(cursor, tableName))

    # cursor.execute("select * from " + tableName)
    # path = os.path.join(exportPath, "table-contents.sql")
    # f = open(path, "a")
    # for row in cursor:
    #     row = rowString(row)
    #     # cur.execute("INSERT INTO " + tableName + " VALUES " + row)
    #     f.write("INSERT INTO " + tableName + " VALUES " + row + "\n")
    # f.close()

def dumpNavPane():
    project.Application.ExportNavigationPane(os.path.join(exportPath, "nav_pane.xml"))

def dumpTables():
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + os.path.abspath(args.project_path) + ';')

    cursor = conn.cursor()
    # tables starting with "_" not included because they were causing errors and they don't have a csv file counterpart?
    tables = [listing[2] for listing in cursor.tables(tableType='TABLE') if listing[2].startswith("_") == False]
    
    count = 1
    for table in tables:
        print("{}/{} tables".format(count, len(tables)), end= "\r")
        dumpTable(table)
        count += 1
    print()

def loadTables():
    filePath = "data.sql"
    with open(filePath) as file:
        lines = len(file.readlines())

    with open(filePath) as file:
        count = 1
        for line in file:
            print("loading tables {}%".format(int(count/lines*100)), end = "\r")
            if line.startswith("CREATE") or line.startswith("INSERT"):
                project.DoCmd.RunSQL(line)
            count += 1
    print()

def loadQueries():
    files = [file for file in os.listdir(exportPath) if file.split(".")[1] == "sql"] 
    
    count = 1
    for file in files:
        print("loading queries {}/{}".format(count, len(files)), end = "\r")
        sql = open(os.path.join(exportPath, file),"r")
        dbName = project.DBEngine.Workspaces(0).Databases(0).Name
        project.DBEngine.Workspaces(0).OpenDatabase(dbName).CreateQueryDef(file.split(".")[0], sql.read())
        count += 1
    print()

def loadForms():
    formNames = [file for file in os.listdir(exportPath) if file.split(".")[1] == "frm"] 

    count = 1
    for formName in formNames:
        print("loading forms {}/{}".format(count, len(formNames)), end = "\r")
        loadForm(formName.split(".")[0])
        count += 1
    print()

def loadModules():
    files = [file for file in os.listdir(exportPath) if file.split(".")[1] == "bas"] 

    count = 1
    for file in files:
        print("loading modules {}/{}".format(count, len(files)), end = "\r")
        moduleName = file.split(".")[0]
        project.Application.LoadFromText(5, moduleName, os.path.join(exportPath, file))
        count += 1
    print()

def loadNavPane():
    project.Application.ImportNavigationPane(os.path.join(exportPath, "nav_pane.xml"))


match args.command:
    case "dump-all":
        dumpAllForms()
        dumpAllModules()
        dumpAllQueries()
        dumpTables()
        
    case "dump-form":
        dumpForm(args.element_name)
    case "load-form":
        loadForm(args.element_name)
    case "dump-module":
        dumpModule(args.element_name)
    case "dump-query":
        dumpQuery(args.element_name)
    case "dump-table":
        dumpTable(args.element_name)
    case "dump-nav-pane":
        dumpNavPane()

    case "dump-forms":
        dumpAllForms()
    case "load-forms":
        loadForms()
    case "dump-modules":
        dumpAllModules()
    case "dump-queries":
        dumpAllQueries()
    case "dump-tables":
        dumpTables()
    
    case "load-tables":
        loadTables()
    case "load-queries":
        loadQueries()
    case "load-modules":
        loadModules()
    case "load-nav-pane":
        loadNavPane()
    
    case "load-all":
        loadTables()
        loadQueries()
        loadForms()
        loadModules()

project.Application.Quit()