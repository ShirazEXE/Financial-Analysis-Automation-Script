import requests
import pandas as pd
import subprocess
from bs4 import BeautifulSoup

URL = 'https://www.moneycontrol.com/financials/'
Company = 'gujaratambujaexports'
CompanyCode = 'GAE'

urls = {}

urls['balance sheet 2018-2022'] = '{}{}/balance-sheetVI/{}#{}'.format(URL,Company,CompanyCode,CompanyCode)
urls['P&L 2018-2022'] = '{}{}/profit-lossVI/{}#{}'.format(URL,Company,CompanyCode,CompanyCode)
urls['balance sheet 2013-2017'] = '{}{}/balance-sheetVI/{}/2#{}'.format(URL,Company,CompanyCode,CompanyCode)
urls['P&L 2013-2017'] = '{}{}/profit-lossVI/{}/2#{}'.format(URL,Company,CompanyCode,CompanyCode)

xlFile = pd.ExcelWriter(f'Financial_Statement.xlsx', engine = 'xlsxwriter')

for key in urls.keys():
    html = requests.get(urls[key])
    soup = BeautifulSoup(html.content, 'html.parser')
    df = pd.read_html(str(soup), attrs = {'class' : 'mctable1'})[0]
    df.to_excel(xlFile, sheet_name = key, index = False)

xlFile.close()
print("Financial statements spreadsheets downloaded!")

subprocess.run('Python3 Fin_Solven_Ratio.py', shell=True)
