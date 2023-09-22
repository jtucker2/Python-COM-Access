import sys
import os
import win32com.client as win32

"""
python kraken.py <project_path> <export_path>
    dump-forms
    dump-form <form_name>
    load-form <form_name>
    load-forms
"""

project = win32.gencache.EnsureDispatch('Access.Application')
project.Application.OpenCurrentDatabase(sys.argv[1])

currentProject = project.Application.CurrentProject
currentData = project.Application.CurrentData

exportPath = sys.argv[2]

def dumpForm(formName):
    project.DoCmd.OpenForm(formName)
    project.Application.SaveAsText(2, formName, exportPath + "\\" + formName + ".frm")
    project.DoCmd.Close(2, formName)

def loadForm(formName):\
    project.Application.LoadFromText(2, formName, exportPath + "\\" + formName + ".frm")

def dumpModule(moduleName):
    project.DoCmd.OpenModule(moduleName)
    project.Application.SaveAsText(5, moduleName, exportPath + "\\" + moduleName + ".bas")
    project.DoCmd.Close(5, moduleName)

def dumpQuery(queryName):
    dbName = project.DBEngine.Workspaces(0).Databases(0).Name
    try:
        queryString = project.DBEngine.Workspaces(0).OpenDatabase(dbName).QueryDefs(queryName).SQL
        path = os.path.join(exportPath, queryName + ".sql")
        f = open(path, "w")
        f.write(queryString)
        f.close()
    except:
        print("Error", queryName)

def extractFileName(fileName):
    return fileName.split(".")[0]

def dumpAllForms():
    allForms = currentProject.AllForms
    formNames = []
    for i in range(allForms.Count):
        formNames.append(allForms.Item(i).Name)

    # TODO: this form wasn't working
    formNames.remove("ProjectVariants")

    for formName in formNames:
        print(formName)
        dumpForm(formName)

def dumpAllModules():
    allModules = currentProject.AllModules
    moduleNames = []
    for i in range(allModules.Count):
        moduleNames.append(allModules.Item(i).Name)

    for moduleName in moduleNames:
        print(moduleName)
        dumpModule(moduleName)

def dumpAllQueries():
    allQueries = currentData.AllQueries
    queryNames = []
    for i in range(allQueries.Count):
        queryNames.append(allQueries.Item(i).Name)
    
    count = 0
    for queryName in queryNames:
        print(str(count) + "/" + str(len(queryNames)), end = "\r")
        dumpQuery(queryName)
        count += 1

match sys.argv[3]:
    case "dump-all":
        dumpAllForms()
        dumpAllModules()
    case "dump-forms":
        dumpAllForms()
    case "dump-form":
        dumpForm(sys.argv[4])
    case "load-form":
        loadForm(sys.argv[4])
    case "load-forms":
        formNames = [extractFileName(name) for name in os.listdir(exportPath)]

        for formName in formNames:
            print(formName)
            loadForm(formName)
    
    case "dump-modules":
        dumpAllModules()
    case "dump-queries":
        dumpAllQueries()
    case "dump-query":
        dumpQuery(sys.argv[4])

project.Application.Quit()