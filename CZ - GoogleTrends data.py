import requests.exceptions
from pytrends.request import TrendReq
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.message import EmailMessage
import ssl
import numpy as np
import pandas as pd
from datetime import datetime
import time
from datetime import date
import pickle

#   Nahrání dat z Excel tabulky
############
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name('google_trend_keys.json', scope)

client = gspread.authorize(creds)

sheet_KW = client.open('Google_trends').worksheet('seznam klíčových slov')
sheet_data_CZ = client.open('Google_trends').worksheet('CZ -data per KW - whole 2022')
sheet_data_SK = client.open('Google_trends').worksheet('SK -data per KW - whole 2022')
sheet_data_HU = client.open('Google_trends').worksheet('HU -data per KW - whole 2022')

sheet_data_CZ_2021_today = client.open('Google_trends').worksheet('CZ -data per KW 2022 - today')


sheet_data_CZ_2022_today = client.open('Google_trends').worksheet('CZ -data per KW 2023 - today')
sheet_data_SK_2022_today = client.open('Google_trends').worksheet('SK -data per KW 2023 - today')
sheet_data_HU_2022_today = client.open('Google_trends').worksheet('HU -data per KW 2023 - today')

list_of_all_sheets = [sheet_data_CZ, sheet_data_SK, sheet_data_HU]
list_of_all_sheets_2021 = [sheet_data_CZ_2022_today, sheet_data_SK_2022_today, sheet_data_HU_2022_today]
list_of_all_sheets_2022 = [sheet_data_CZ_2022_today, sheet_data_SK_2022_today, sheet_data_HU_2022_today]
list_of_all_sheets_comparison = [list_of_all_sheets_2021, list_of_all_sheets_2022]

email_addresses = sheet_KW.col_values(1)[1:]
CZ_keywords = sheet_KW.col_values(3)[1:]
SK_keywords = sheet_KW.col_values(5)[1:]
HU_keywords = sheet_KW.col_values(7)[1:]
list_of_all_KW = [CZ_keywords, SK_keywords, HU_keywords]

# Již doplněná slova
CZ_keywords_done = sheet_data_CZ.col_values(2)[1:]
SK_keywords_done = sheet_data_SK.col_values(2)[1:]
HU_keywords_done = sheet_data_HU.col_values(2)[1:]
list_of_all_done_KW = [CZ_keywords_done, SK_keywords_done, HU_keywords_done]

###############
pytrends = TrendReq()
#pytrends = TrendReq(
#    hl='da')  # , tz=360, timeout=(10,25), retries=2, backoff_factor=0.1, requests_args={'verify':False})
keywords = []

filename = 'message'
infile = open(filename,'rb')
outfile = open(filename,'wb')

CZ_trending_keyword_100 = []
CZ_trending_keyword_50 = []
CZ_trending_keyword_25 = []

SK_trending_keyword_100 = []
SK_trending_keyword_50 = []
SK_trending_keyword_25 = []

HU_trending_keyword_100 = []
HU_trending_keyword_50 = []
HU_trending_keyword_25 = []

trending_keyword_100 = [CZ_trending_keyword_100, SK_trending_keyword_100, HU_trending_keyword_100]
trending_keyword_50 = [CZ_trending_keyword_50, SK_trending_keyword_50, HU_trending_keyword_50]
trending_keyword_25 = [CZ_trending_keyword_25, SK_trending_keyword_25, HU_trending_keyword_25]

no_data_keywords = []

cat = '0'
#geo = ['CZ', 'SK', 'HU']

today = date.today()
year, month, day = str(today).split('-')
last_whole_year = str(int(year) - 1) + '-' + '01-01' + ' ' + str(int(year) - 1) + '-' + '12-31'
last_year = str(int(year) - 1) + '-' + '01-01' + ' ' + str(int(year) - 1) + '-' + month + '-' + day
this_year = str(year) + '-' + '01-01' + ' ' + str(today)

timeframes = ['today 5-y', 'today 12-m', 'today 3-m', 'today 1-m',
              last_whole_year, this_year, last_year]
error_cnt = 0

# Email sender
def email(text, email_address):
    email_sender = 'alza.emailsender@gmail.com'  # heslo je Alza1234
    email_pass = 'fdevnapfselqfkvn'
    email_receiver = email_address

    subject = 'Google Trends report'
    body = text

    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = subject
    em.set_content(body)
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_pass)
        smtp.sendmail(email_sender, email_receiver, em.as_string())

# Error count
def error_count():
    global error_cnt
    error_cnt += 1
    return error_cnt

# Smazání předchozích hodnot
def delete_data(country):
    country.clear()
    print('Smazano: ', country)


# Google trends pro rok 2021
def google_trends_last_year(timeframe):
    try:
        pytrends.build_payload(keywords,
                               cat,
                               timeframe,
                               geo = 'CZ',
                               gprop='')
        data = pytrends.interest_over_time()
        if data.empty == False:
            #   Vložení dat do sheetu - porovnání
            data['datetime'] = data.index
            data = data.astype({"datetime": str}, errors='raise')
            data = data[['datetime', kw]]
            data.insert(1, 'KW', kw)
            rows_list = []
            for index, row in data.iterrows():
                row = [row[0], kw, row[2]]
                rows_list.append(row)
            list_of_all_sheets[i].insert_rows(rows_list, 1)
            print(kw, ' - hotovo')
            time.sleep(2)
        else:
            no_data_keywords.append(kw)
            print(kw, ' nejsou data.')
            print(no_data_keywords)


    except Exception as error:
        print('Celý rok 2022 - ', error)
        time.sleep(5)
        google_trends_last_year(timeframes[4])


# Porovnání tohoto a minulého roku
def google_trends_comparison():
    try:
        #   Letos
        pytrends.build_payload(keywords,
                               cat,
                               timeframes[5],
                               geo = 'CZ',
                               gprop='')

        data_1 = pytrends.interest_over_time()
        #mean_1 = round(data_1.mean(), 2)
        week_1 = round(data_1[kw][-1:].mean(), 1)  # průměr hledání za 1 týden

        time.sleep(60)

        # Pred rokem
        pytrends.build_payload(keywords,
                               cat,
                               timeframes[6],
                               geo= 'CZ',
                               gprop='')
        #       tvorba proměných
        data_2 = pytrends.interest_over_time()
        #mean_2 = round(data_2.mean(), 2)  # průměr hledání za čtvrt roku
        week_2 = round(data_2[kw][-1:].mean(), 1)  # průměr hledání za 1 týden

        #   Tvorba reportu
        if week_1 and week_2 != 0.0:
            trend = round(((week_1 / week_2) - 1) * 100, 2)
            message_comp = str(trend) + ' % YoY - (loňský týden ' + str(week_2) + ', letošní týden = ' + str(week_1) + ')'
        else:
            message_comp = '- nejsou data'

        #   Vložení dat do sheetu - porovnání
        data = data_1
        data['datetime'] = data.index
        data = data.astype({"datetime": str}, errors='raise')
        data = data[['datetime', kw]]
        data.insert(1, 'KW', kw)
        rows_list = []
        for index, row in data.iterrows():
            row = [row[0], kw, row[2]]
            rows_list.append(row)
        list_of_all_sheets_2022[i].insert_rows(rows_list, 1)
        time.sleep(60)

        #   Vložení dat do sheetu - porovnání
        data = data_2
        data['datetime'] = data.index
        data = data.astype({"datetime": str}, errors='raise')
        data = data[['datetime', kw]]
        data.insert(1, 'KW', kw)
        rows_list = []
        for index, row in data.iterrows():
            row = [row[0], kw, row[2]]
            rows_list.append(row)
        list_of_all_sheets_2021[i].insert_rows(rows_list, 1)

        return message_comp


    except Exception as error:
        error_cnt = error_count()
        if error_cnt > 10:
            error_cnt
            time.sleep(100)
        print('Porovnání loňského a letošního roku - ', error)
#        time.sleep(3)
        google_trends_comparison()


# Google trends za časový úsek
def google_trends():
    global trend_7_d
    try:
        pytrends.build_payload(keywords,
                               cat,
                               timeframes[2],
                               geo = 'CZ',
                               gprop='')
        #   Trend hledanosti porovnání 3 měsíce / 7 dní
        data = pytrends.interest_over_time()
        mean = round(data.mean(), 2)  # průměr hledání za čtvrt roku
        mean_7_d = round(data[kw][-7:].mean(), 2)  # průměr hledání za 1 týden
        trend_7_d = round(((mean_7_d / mean[kw]) - 1) * 100, 1)
        #   Porovnani s minulim rokem
        message_comp = google_trends_comparison()

        if len(kw) < 6:
            message = "'" + kw + "'" + '\t\t\t' + str(trend_7_d) + ' % \t\t\t' + message_comp
        elif len(kw) > 14:
            message = "'" + kw + "'" + '\t' + str(trend_7_d) + ' % \t\t\t' + message_comp
        else:
            message = "'" + kw + "'" + '\t\t' + str(trend_7_d) + ' % \t\t\t' + message_comp


        #  Rozřazení podle kategorie
        if trend_7_d > 100:
            trending_keyword_100[i].append(message)
        elif trend_7_d > 50:
            trending_keyword_50[i].append(message)
        elif trend_7_d > 25:
            trending_keyword_25[i].append(message)
        print(kw, ' - hotovo.')


        pickle.dump(message, outfile)

    #   Chyba scriptu
    except Exception as error:
        error_cnt = error_count()
        if error_cnt > 10:
            error_cnt
            time.sleep(100)
        print('Categorization module - ', error)
 #       time.sleep(30)
        google_trends()


###     SAMOTNY PRŮBĚH APLIKACE     ###


# Nahrání dat za předchozí rok
print('Kontroluji nově přidaná klíčová slava a přidávám je do databáze.. ')
i = 0
for kw in CZ_keywords:
    if kw not in CZ_keywords_done:
        keywords.append(kw)
        google_trends_last_year(timeframes[4])
        keywords.pop()
        time.sleep(60)
i += 1
print(10 * '*', ' Hotovo ', 10 * '*')

#   Hledanost KW za poslední 3 měsíce + YoY + tvorba reportu

# Smazání předchozích hodnot

delete_data(sheet_data_CZ_2021_today)
delete_data(sheet_data_CZ_2022_today)


# hledanost a porovnani
print('Kategorizuji klíčová slova a vytvářím report..')
i = 0                                                               # 0=CZ  1=SK    2=HU
errors_count = 0
for kw in CZ_keywords:
    keywords.append(kw)
    google_trends()
    keywords.pop()
i += 1
print(10 * '*', ' Hotovo ', 10 * '*')

# Spojení listů na string
CZ_trending_keyword_100 = '\n'.join(CZ_trending_keyword_100)
CZ_trending_keyword_50 = '\n'.join(CZ_trending_keyword_50)
CZ_trending_keyword_25 = '\n'.join(CZ_trending_keyword_25)
SK_trending_keyword_100 = '\n'.join(SK_trending_keyword_100)
SK_trending_keyword_50 = '\n'.join(SK_trending_keyword_50)
SK_trending_keyword_25 = '\n'.join(SK_trending_keyword_25)
HU_trending_keyword_100 = '\n'.join(HU_trending_keyword_100)
HU_trending_keyword_50 = '\n'.join(HU_trending_keyword_50)
HU_trending_keyword_25 = '\n'.join(HU_trending_keyword_25)

#   Odeslání emailu s upozorněním pro jednotlive zeme
pre_text = 'Ahoj, \n \n zasílám týdenní přehled růstu hledanosti klíčových slov.\n Jde o porovnání posledních 7 dní s předchozími 3 měsíci a \n' + \
           ' procentuální růst/pokles při porovnání posledních 7 dnů YoY.\n\n'

CZ_text = 'CZ report,\n  \n' + \
          'Nárůst o 100 % a více: \n' + str(CZ_trending_keyword_100) + '\n' + ' \n' + \
          'Nárůst o 50 % a více: \n' + str(CZ_trending_keyword_50) + '\n' + ' \n' + \
          'Nárůst o 25 % a více: \n' + str(CZ_trending_keyword_25) + '\n' + ' \n' + \
          ' Odkaz na DataStudio: https://datastudio.google.com/reporting/99b277f0-9a74-4998-b5a8-26442169f35e/page/t8O5C '

email_text = pre_text + CZ_text + '\n\n'

#   Poslani emailu
for email_address in email_addresses:
    email(email_text, email_address)

print('Emaily odeslány.')

#print(no_data_keywords)
