import os
import shutil

#import Test
#from requests import get
#from csv import DictReader
#from bs4 import BeautifulSoup as Soup
#from io import StringIO
#import pandas
import subprocess
import datetime
import openpyxl

print('Buffer Selected')

		uploadForm2  = FORM(

		    LABEL("* ", "Upload Updated Buffer File ", " *", BR(), INPUT(_type='file', _name="UploadBufferName", requires=IS_NOT_EMPTY()), _class='btn btn-primary btn-lg btn-file', _style='font-size:17px; background-color:#33446e;'),
		    BR(),
		    
		    HR(),
		    BR(),
		    
		    HR(),
		    INPUT(_type='submit', _value='RUN', _class='btn btn-warning btn-lg col-md-2', _style="background-color:#33446e;font-size:14px;"),
		    #HR(),
		    #INPUT(_type='file', _value='Download Output', _class='btn btn-warning', _style="background-color:#33446e;font-size:14px;"),
		    HR()
		    #INPUT(SELECT([OPTION(i, _value=str(i)) for i in Destination],  BR(), _class="form-control",_type='checkbox' _name='fileSelect3',_align="center", _value='Select a file'), BR(), _class ="col-md-4", _style="font-size:17px;"),
			)
		if uploadForm2.process(formname='af').accepted:
		    uploadStatus="Form is being validated."
		    
		    #filename = str(request.vars.fileSelect1) + "_" + str(request.vars.fileSelect2) + ".xlsx"
		    filename = "Buffer" + ".xlsx"
		    print(filename)

		    upload_folder_temp = "C:\\web2py_win\\web2py\\applications\\IndiaPost_UI\\Source Code\\IndiaPost\\"
		    file = request.vars.UploadBufferName.file
		    path = upload_folder_temp +"\\" + filename
		    print(file)

		if not(os.path.exists(upload_folder_temp)):
		    os.makedirs(upload_folder_temp)
		    uploadStatus="New folder is created."

		shutil.copyfileobj(file, open(path,'wb'))