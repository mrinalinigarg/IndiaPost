import csv
from datetime import datetime, timedelta
import openpyxl
import pandas as pd

path = "C:\\web2py_win\\web2py\\applications\\IndiaPost_UI\\Source Code\\IndiaPost\\Test.xlsx"


curDir = "C:\\web2py_win\\web2py\\applications\\IndiaPost_UI\\input files\\"
df_loc = pd.read_excel(path) #Reading the Values of 369 locations
loc_id_list = df_loc.values.tolist() #Creating a list loc_id_list for 369 locations
num_loc = len(loc_id_list)
#print(num_loc)

print(len(df_loc.columns))

if(len(df_loc.columns)!=13):
	print("Error in File. Please check.")
else:
	print("File Accepted")