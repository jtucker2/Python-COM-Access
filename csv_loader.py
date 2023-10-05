import sys
import os
import sqlite3
import pandas as pd

def loadCSV(path, tableName):
    con = sqlite3.connect("DomainModel.db", isolation_level=None)
    csv = pd.read_csv(path)
    csv.to_sql(tableName, con, if_exists='append', index = False)

def loadCSVs(path):
    files = os.listdir(path)
    count = 1
    for file in files:
        print("{}/{} tables".format(count, len(files)), end= "\r")
        loadCSV(os.path.join(path, file), file.split(".")[0])
        count += 1
    print()

loadCSVs(sys.argv[1])