# import webbrowser

# url = 'http://obecimso.net/web63/listOfficeRoomR1.php'

# #Windows
# chrome_path = 'C:/Users/nattawat.kan/AppData/Local/Google/Chrome/Application/chrome.exe %s'

# webbrowser.get(chrome_path).open(url)


import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
  
# The webpage URL whose table we want to extract 
url = "http://obecimso.net/web63/listStudentRegisR1.php?area_code=013"
  
# Read data must use array[1]
table = pd.read_html(url)[1] 
 
writer = ExcelWriter('D:/python_sc/dataFromWeb.xlsx')
table.to_excel(writer, 'Sheet1', index=False)
writer.save()


