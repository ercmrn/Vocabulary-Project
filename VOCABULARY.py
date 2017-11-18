import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import lxml.etree
import sys
import os
import pandas as pd

# non_bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)
# I suspect this will come up later. Not all characters are handled well.

def main():
    wks = access_worksheet()
    jmdict_xml, expr = load_jmdict()
    df = get_data(wks, jmdict_xml, expr)
    update_sheet(df, wks)


def access_worksheet():
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'C:/Users/Eric/AppData/Local/Programs/Python/Python36/My Project-cc4d575f35d9.json', scope)
    gc = gspread.authorize(credentials)
    wks = gc.open("VOCABULARY").sheet1

    return wks


def load_jmdict():
    jmdict_filepath = 'C:/Users/Eric/Downloads/JMdict_e/JMdict_e.xml'
    jmdict_xml = lxml.etree.parse(jmdict_filepath)
    expr = '//keb[text()="{}"]/../..//{}'
    return jmdict_xml, expr


def extract(jmdict_xml, expr, key, element):
    return [x.text for x in jmdict_xml.xpath(expr.format(key, element))]


def list_trim(inputlist, max_length):
    if len(inputlist) > max_length:
        return inputlist[0:max_length]
    else:
        return inputlist


def get_data(worksheet, jmdict_xml, expr):
    testdf = pd.DataFrame(columns=['row', 'word', 'reading', 'meaning', 'pos', 'date'])
    rows, missing_words, readings, meanings, pos, dates = [], [], [], [], [], []
    num_words = len([x for x in worksheet.col_values(1) if x != '']) + 1  # because I don't want to loop the whole sheet...?

    for c in range(2, num_words):
        if (worksheet.cell(c, 1).value != "") & (worksheet.cell(c, 3).value == ""):
            rows.append(c)
            missing_words.append(worksheet.cell(c, 1).value)
        else:
            continue

    testdf['row'] = rows
    testdf['word'] = missing_words

    for word in testdf['word']:
        readings.append(list_trim(extract(jmdict_xml, expr, word, 'reb'), 8))
        meanings.append(list_trim(extract(jmdict_xml, expr, word, 'gloss'), 3))
        pos.append(list_trim(extract(jmdict_xml, expr, word, 'pos'), 3))
        dates.append(datetime.date.today())

    testdf['reading'] = readings
    testdf['meaning'] = meanings
    testdf['pos'] = pos
    testdf['date'] = dates

    return testdf

def update_sheet(testdf, wks):
    for i, x in enumerate(testdf['row']):
        wks.update_cell(x, 3, testdf['reading'][i])
        wks.update_cell(x, 4, testdf['meaning'][i])
        wks.update_cell(x, 5, testdf['pos'][i])
        wks.update_cell(x, 6, testdf['date'][i])

if __name__ == '__main__':
    main()

def lambda_handler():
    main()
