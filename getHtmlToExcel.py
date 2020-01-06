# import webbrowser

# url = 'http://obecimso.net/web63/listOfficeRoomR1.php'

# #Windows
# chrome_path = 'C:/Users/nattawat.kan/AppData/Local/Google/Chrome/Application/chrome.exe %s'

# webbrowser.get(chrome_path).open(url)


import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

mainURL = 'http://obecimso.net/web63/listOfficeRoomR1.php'

# Need for loop this URL 
dataURL = "http://obecimso.net/web63/listStudentRegisR1.php?area_code=013"
  
# Read data must use array[1]
table = pd.read_html(dataURL)[1] 
 
# Need re-name excel with label from Website link (use variable with array insted fix String) 
writer = ExcelWriter('D:/python_sc/Excel/dataFromWeb.xlsx')
table.to_excel(writer, 'Sheet1', index=False)
writer.save()


