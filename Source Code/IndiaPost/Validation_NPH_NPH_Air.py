import csv
from datetime import datetime, timedelta
import openpyxl
import time
import math
import os
import pandas as pd

def main():

	filename = 'Air_NPH to NPH.xlsx'

	list_NPH = ['VIJAYAWADA','VISAKHAPATNAM','GUWAHATI','PATNA','MUZAFFARPUR','RAIPUR','DELHI','AHMEDABAD','SURAT','VADODARA','AMBALA','FARIDABAD','GURGAON','SHIMLA','PATHANKOT','JAMMU','SRINAGAR','JAMSHEDPUR','RANCHI','BANGALORE','HUBLI','MANGALORE','MYSORE','KOCHI','KOZHIKODE','TRIVANDRUM','BHOPAL','INDORE','AURANGABAD','MUMBAI','NAGPUR','PANAJI','PUNE','AGARTALA','AIZAWL','DIMAPUR','IMPHAL','BHUBANESWAR','CHANDIGARH','JALANDHAR','LUDHIANA','JAIPUR','CHENNAI','COIMBATORE','MADURAI','TRICHY','HYDERABAD','AGRA','ALLAHABAD','GHAZIABAD','LUCKNOW','BAREILLY','DEHRADUN','KOLKATA','SILIGURI','56 APO','99 APO']
	#print(list_NPH[4])
	list_days_week = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

	df_loc = pd.read_excel("C:\\web2py_win\\web2py\\applications\\IndiaPost_UI\\Source Code\\IndiaPost\\"+filename) 
	loc_id_list = df_loc.values.tolist() 
	num_loc = len(loc_id_list)

	valid_counter_o = ""
	valid_counter_d = ""
	valid_counter_t = ""
	valid_counter_w = ""
	time1=""
	time2=""
	time3=""
	counter = 0
	for i in range(num_loc):
		origin = loc_id_list[i][1]
		destination = loc_id_list[i][2]
		time = loc_id_list[i][3]
		week = loc_id_list[i][4]
		#print(origin,destination)
		try:
			time1 = int(time[0:2])
			time2 = str(time[2])
			time3 = int(time[3:5])
		except:
			print('An error occurred in Time Format.')
			return("An error occurred in Time Format.")

		#print(week)
		#try:
		if(origin in list_NPH):
			pass
			#print("TRUE")
			#valid_counter_o = "Origin Input is Correct"
		else:
			#print("FALSE")
			valid_counter_o = "Origin in not an NPH."
			counter += 1
		if(destination in list_NPH):
			pass
			#print("TRUE")
			#valid_counter_d = "Destination Input is Correct"
		else:
			#print("FALSE")
			valid_counter_d = "Destination in not an NPH."
			counter += 1
		if(time1<=24) and (time2==":") and (time3<=60):
			pass
			#print(time1,time2,time3)
			#valid_counter_t = "Time Input is Correct"
		else:
			#print("incorrect input time format, please enter as a string 17:30")
			valid_counter_t = "Incorrect time format, please enter as a string 17:30."
			counter += 1
		if(week in list_days_week):
			pass
			#valid_counter_w = "Weekday Input is Correct"
		else:
			#print("FALSE")
			valid_counter_w = "Weekday Input is incorrect, please mention weekday as 'Mon', 'Tue', etc"
			counter += 1

	if(counter==0):
		print("Form Accepted.Code is executing.")
		return("Form Accepted.Code is executing.")
	else:
		print('Number of Incorrect Cell Entries is: '+str(counter)+'.',valid_counter_o,valid_counter_d,valid_counter_t,valid_counter_w)
		return('Number of Incorrect Cell Entries is: '+str(counter)+'.', valid_counter_o,valid_counter_d,valid_counter_t,valid_counter_w)


if __name__ == '__main__':
	main()