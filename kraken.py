import sys

import win32com.client as win32

project = win32.gencache.EnsureDispatch('Access.Application')
project.Application.OpenCurrentDatabase("E:\Documents\Database2.accdb")

currentProject = project.Application.CurrentProject
currentData = project.Application.CurrentData

exportPath = sys.argv[1]

def dumpForm(formName):
    project.DoCmd.OpenForm(formName)
    project.Application.SaveAsText(2, formName, exportPath + "\export_" + formName + ".txt")
    project.DoCmd.Close(2, formName)

match sys.argv[2]:
    case "dump-forms":
        allForms = currentProject.AllForms
        formNames = []
        for i in range(allForms.Count):
            formNames.append(allForms.Item(i).Name)
        for formName in formNames:
            dumpForm(formName)

    case "dump-form":
        dumpForm(sys.argv[3])

project.Application.Quit()