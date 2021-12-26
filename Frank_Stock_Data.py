import gspread
from datetime import datetime
from datetime import timedelta
import time
from oauth2client.service_account import ServiceAccountCredentials
import yahoo_fin.stock_info as si
import requests
from bs4 import BeautifulSoup
import pandas as pd
import lxml


def datePlusOne(str):
    date = datetime.strptime(str, '%b %d, %Y')
    addOne = timedelta(days = 1)
    date = date + addOne
    strDate = date.strftime('%b %d, %Y')
    return strDate

def is_float(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def UpdateCellFunc(spreadsheet, row, column, value):
	while True:
		try:
			spreadsheet.update_cell(row, column, value)
			time.sleep(1)
			break
		except:
			time.sleep(1)

scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds =  ServiceAccountCredentials.from_json_keyfile_name('/home/pi/Documents/Python/google_creds.json', scope)

client = gspread.authorize(creds)

sheet = client.open('Portfolio').sheet1

rowj = 0

for j in range(2, 40):
    time.sleep(1.01)
    rowj = j
    if(sheet.cell(j, 1).value == None):
        break

rowi = 0

for i in range(2, rowj):
    rowi = i
    ticker = sheet.cell(i, 1).value
    data = si.get_stats(ticker)
    time.sleep(1)
    ex_date = data['Value'][26]
    UpdateCellFunc(sheet, i, 22, datePlusOne(ex_date))
    #sheet.update_cell(i, 22, datePlusOne(ex_date))
    #time.sleep(1)
    pay_date = data['Value'][25]
    UpdateCellFunc(sheet, i, 23, datePlusOne(pay_date))
    #sheet.update_cell(i, 23, datePlusOne(pay_date))
    #time.sleep(1)
    div = data['Value'][19]
    if isinstance(div, float) : UpdateCellFunc(sheet, i, 10, "")
        #sheet.update_cell(i, 9, "")
    else : UpdateCellFunc(sheet, i, 10, div)
        #sheet.update_cell(i, 10, div)
    #time.sleep(1)
    beta = data['Value'][0]
    if isinstance(beta, float) : UpdateCellFunc(sheet, i, 24, "")
        #sheet.update_cell(i, 24, "")
    else : UpdateCellFunc(sheet, i, 24, beta)
        #sheet.update_cell(i, 24, beta)
    #time.sleep(1)
    # payout_ratio = data['Value'][24]
    # print(type(payout_ratio))
    # if isinstance(payout_ratio, float) : sheet.update_cell(i, 28, "")
    # else : sheet.update_cell(i, 28, payout_ratio)


    URL = "https://www.reuters.com/companies/" + ticker + ".N/key-metrics"
    page = requests.get(URL)
    soup = BeautifulSoup(page.content, 'html.parser')
    nums = soup.find_all('div', class_='KeyMetrics-table-container-3wVZN')
    num = nums[6]
    cagr = num.find_all('td')
    strCagr = cagr[5].get_text()
    if(is_float(strCagr)): numCagr = float(strCagr) / 100
    else: numCagr = strCagr
    #sheet.update_cell(i, 26, numCagr)
    UpdateCellFunc(sheet, i, 26, numCagr)
    #time.sleep(1)

now = datetime.now()
#minus_five = timedelta(hours = 5)
#now_EST = now - minus_five
UpdateCellFunc(sheet, rowi + 6, 3, now.strftime('%H:%M:%S %b %d, %Y'))
#sheet.update_cell(rowi + 6, 3, now.strftime('%H:%M:%S %b %d, %Y'))
