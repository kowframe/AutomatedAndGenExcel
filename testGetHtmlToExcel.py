# import webbrowser

# url = 'http://obecimso.net/web63/listOfficeRoomR1.php'

# #Windows
# chrome_path = 'C:/Users/nattawat.kan/AppData/Local/Google/Chrome/Application/chrome.exe %s'

# webbrowser.get(chrome_path).open(url)

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import requests
from bs4 import BeautifulSoup
import urllib.request
import re

race_link = 'http://obecimso.net/web63/listOfficeRoomR1.php'
sauce1 = urllib.request.urlopen(race_link).read()
soup1 = BeautifulSoup(sauce1, 'html.parser')
pathExcelOutPut = 'D:/python_sc/testExcel/'
areaName = 0
areaNameStr = ''

for link in soup1.find_all('tr'):
	for ilink in link.find_all('td'):
		#print(ilink.get_text())
		#print(areaName)
		#areaName += 1
		print(areaNameStr)
		for jlink in ilink.find_all(href=re.compile("listStudent")):
			areaName += 1
			areaNameStr = str(areaName)
			print(jlink.get('href'))

print('Done!')
