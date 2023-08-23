from pytrends.request import TrendReq
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.message import EmailMessage
import ssl
import numpy as np
#import matplotlib.pyplot as plt
from datetime import datetime
import time

keyword = []

    ###   Nahrání dat z Excel tabulky   ###
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('google_trend_keys.json',scope)
client = gspread.authorize(creds)

# Definování jednotlivých google sheetů a sloupců
sheet_KW = client.open('Google_trends').worksheet('seznam klíčových slov')
sheet_data = client.open('Google_trends').worksheet('data per KW')

# Sloupce
CZ_keywords = sheet_KW.col_values(3)[1:]
SK_keywords = sheet_KW.col_values(4)[1:]
HU_keywords = sheet_KW.col_values(5)[1:]
email_adresses = sheet_KW.col_values(1)[1:]

kw = ['máslo', 'mléko']

pytrends = TrendReq(hl='da')
#timeframe = ['today 12-m','today 3-m'] mozne volby

def google_trends():
    pytrends.build_payload( kw_list= kw,                                              # klicova slova
                            cat=0,                                          # kategorie
                            timeframe='2021-01-01 2021-12-31',              # za jaký časový úsek
                            geo='CZ',                                       # misto, ceska republika =CZ
                            gprop='')                                       # z jakého zdroje je brát input (news, images, youtube.

            #   tvorba proměných
    data = pytrends.interest_over_time()
    print(data)
    return data

for kw in kw:
    keyword.append(kw)
    google_trends()
    keyword.pop()

    #sheet_data.insert_row(data, 2)




