import requests
from bs4 import BeautifulSoup
import pandas as pd
import progressbar
import re
from datetime import datetime
import pygsheets
import csv
# set Google Sheets variables
import user_specific_variables

# variable
CONFIG_SYMBOLS = []
CONFIG_COLUMN_SYMBOL_IX = 0

def scrape_finviz(symbols):

    # Get Column Header
    hdr = {'User-Agent': 'Mozilla/5.0'}
    req = requests.get("https://finviz.com/quote.ashx?t=FB",headers=hdr)
    if req.status_code == 403:
        print("Error 403")
        exit(1)
    soup = BeautifulSoup(req.content, 'html.parser')
    table = soup.find_all(lambda tag: tag.name == 'table')
    rows = table[7].findAll(lambda tag: tag.name == 'tr')
    out = []
    for i in range(len(rows)):
        td = rows[i].find_all('td')
        out = out + [x.text for x in td]

    # Adding actual headers: Ticker, Sector, Sub-sector and country plus all from the finviz table
    # and then 5 from GuruFocus
    guru_ls = ['Piotroski F-Score', 'Altman Z-Score', 'Beneish M-Score']
    ls = ['Ticker', 'Sector', 'Sub-Sector', 'Country'] + out[::2] + guru_ls + ['ROIC', 'WACC']

    dict_ls = {k: ls[k] for k in range(len(ls))}
    df = pd.DataFrame()
    p = progressbar.ProgressBar()
    p.start()
    for j in range(len(symbols)):
        p.update(j / len(symbols) * 100)

        # Initialize FinViz parsing
        req = requests.get("https://finviz.com/quote.ashx?t=" + symbols[j],headers=hdr)
        if req.status_code != 200:
            continue
        soup = BeautifulSoup(req.content, 'html.parser')
        table = soup.find_all(lambda tag: tag.name == 'table')

        # Initialize GuruFocus table parsing (3 values for all the symbols)
        if symbols[j].find('-'):
            guru_symbol = symbols[j].replace('-', '.')
        guru_req = requests.get("https://www.gurufocus.com/stock/" + guru_symbol,headers=hdr)
        if guru_req.status_code != 200 and guru_req.status_code != 403:
            continue
        guru_soup = BeautifulSoup(guru_req.content, 'html.parser')

        # Process tables from BS
        rows = table[5].findAll(lambda tag: tag.name == 'tr')
        sector = []
        for i in range(len(rows)):
            td = rows[i].find_all('td')
            sector = sector + [x.text for x in td]
        if(sector):
            sector = sector[2].split('|')
        rows = table[7].findAll(lambda tag: tag.name == 'tr')
        out = []
        for i in range(len(rows)):
            td = rows[i].find_all('td')
            out = out + [x.text for x in td]
        out = [symbols[j]] + sector + out[1::2]

        out_df = pd.DataFrame(out).transpose()
        df = df.append(out_df, ignore_index=True)

        scores = []

        for val in guru_ls[0:]:
            try:
                scores.append(guru_soup.find('a', string=re.compile(val)).find_next('td').text)
            except:
                scores.append('')

        try:
            roic_value = re.search(r'ROIC \d+\.\d+', guru_soup.getText()).group(0)
            scores.append(roic_value.replace('ROIC ', ''))
            wacc_value = re.search(r'WACC \d+\.\d+', guru_soup.getText()).group(0)
            scores.append(wacc_value.replace('WACC ', ''))
        except:
            scores.append('')
            scores.append('')

        df_len = len(df) - 1

        df.loc[df_len, guru_ls[0]] = scores[0]
        df.loc[df_len, guru_ls[1]] = scores[1]
        df.loc[df_len, guru_ls[2]] = scores[2]
        df.loc[df_len, ls[-2]] = scores[3]
        df.loc[df_len, ls[-1]] = scores[4]

    p.finish()
    df = df.rename(columns=dict_ls)

    gc = pygsheets.authorize(service_file=user_specific_variables.json_file)
    sheet = gc.open_by_key(user_specific_variables.sheet_key)

    worksheet = sheet.worksheet_by_title(user_specific_variables.worksheet_title)

    worksheet.clear(start='A1')
    worksheet.set_dataframe(df, start='A1', nan='')

    # Write output CSV from dataframe as a backup to local working directory as outputYYYY-MM-DD.csv
    output_file_with_date = 'output' + datetime.today().strftime('%Y-%m-%d') +'.csv'
    df.to_csv(output_file_with_date, index=False)

    return (df)

#data = scrape_finviz(['msft', 'fb', 'aapl'])

with open('tickers.csv', newline='') as csvfile:
    reader = csv.reader(csvfile, delimiter=',', quotechar='"')
    linenumber = 1
    for lineContent in reader:
        if linenumber > 1:
            CONFIG_SYMBOLS.append(lineContent[CONFIG_COLUMN_SYMBOL_IX])
        linenumber = linenumber + 1
    print(CONFIG_SYMBOLS)

data = scrape_finviz(CONFIG_SYMBOLS)

