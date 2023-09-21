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
    project.Application.SaveAsText(2, formName, exportPath + "\export_" + formName + ".txt")
    project.DoCmd.Close(2, formName)

def loadForm(formName):
    project.Application.LoadFromText(2, formName, exportPath + "\export_" + formName + ".txt")

def extractFileName(fileName):
    return fileName[7:].split(".")[0]

match sys.argv[3]:
    case "dump-forms":
        allForms = currentProject.AllForms
        formNames = []
        for i in range(allForms.Count):
            formNames.append(allForms.Item(i).Name)

        # TODO: this form wasn't working
        formNames.remove("ProjectVariants")

        for formName in formNames:
            print(formName)
            dumpForm(formName)

    case "dump-form":
        dumpForm(sys.argv[4])
    
    case "load-form":
        loadForm(sys.argv[4])

    case "load-forms":
        formNames = [extractFileName(name) for name in os.listdir(exportPath)]

        for formName in formNames:
            print(formName)
            loadForm(formName)

project.Application.Quit()