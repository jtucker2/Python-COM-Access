import win32com.client as win32
import pandas as pd
from pathlib import Path

project = win32.gencache.EnsureDispatch('Access.Application')
project.Application.OpenCurrentDatabase("E:\Documents\Database2.accdb")

exportPath = "E:\Documents"

# ADODB = win32.Dispatch("ADODB.Connection")

# fields = list(access._prop_map_get_.keys())
# print(fields)

currentProject = project.Application.CurrentProject
currentData = project.Application.CurrentData

# print(project.ADOConnectString)
# print(project.COMAddIns)

# ADODB.OpenRecordset()


# TABLE INFOMRATION
# tables = currentData.AllTables
# print(tables.Count)
# print(tables.Item(tables.Count-1))
# print(project.DBEngine.Workspaces.Count)
# # print(project.DBEngine.Workspaces(0).TableDefs.Count)
# print(project.DBEngine.Workspaces(0).Name)
# print(project.Application.TableDefs.Count)


# properties = project.DBEngine.Properties
# for i in range(properties.count):
#     print(properties(i).Name)


# COUNT CONTROLS
allForms = currentProject.AllForms
formNames = []
for i in range(allForms.Count):
    formNames.append(allForms.Item(i).Name)

controlCount = 0
for formName in formNames:
    project.DoCmd.OpenForm(formName)
    form = project.Forms.Item(0)
    controlCount += form.Controls.Count
    project.Application.SaveAsText(2, formName, exportPath + "\export_" + formName + ".txt")
    project.DoCmd.Close(2, formName)


print("The project has", allForms.Count, "forms and defines a total of", controlCount, "controls")


# project.Visible = True

# _ = input("Press ENTER to quit:")

project.Application.Quit()