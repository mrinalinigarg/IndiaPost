import csv
from datetime import datetime, timedelta
import openpyxl
import pandas as pd

#curDir = "C:\\web2py_win\\web2py\\applications\\IndiaPost_UI\\input files\\"
curDir = "C:\\web2py_win\\web2py\\applications\\IndiaPost_UI\\input files\\"
pwd = "C:\\web2py_win\\web2py\\applications\\IndiaPost_UI\\Source Code\\IndiaPost\\"
df_loc = pd.read_excel(curDir+'Location_Details.xlsx') #Reading the Values of 369 locations
loc_id_list = df_loc.values.tolist() #Creating a list loc_id_list for 369 locations
num_loc = len(loc_id_list)
print(num_loc)

data_list = [[] for x in range(num_loc)]

for i in range(num_loc):
	data_list[i].append(loc_id_list[i][1])
	data_list[i].append(loc_id_list[i][0] - 1)

with open(pwd+'loc_to_id.csv', 'w', newline = '') as fp_out: #Writing the values as list of lists in 1 go to decrease time complexity
	writer = csv.writer(fp_out)
	writer.writerows(data_list)