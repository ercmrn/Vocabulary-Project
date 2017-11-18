import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import lxml.etree
import sys
import os
from collections import defaultdict

scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('My Project-cc4d575f35d9.json', scope)
gc = gspread.authorize(credentials)
wks = gc.open("VOCABULARY").sheet1

def new_words():
    num_words = len([x for x in wks.col_values(1) if x != '']) + 1
    missing_words = defaultdict()

    for c in range(2, num_words):
        if (wks.cell(c, 1).value != "") & (wks.cell(c, 3).value == ""):
            missing_words[(wks.cell(c, 1).value)] = []
        else:
            continue
        
    return missing_words

non_bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)

def load_jmdict():
    jmdict_filepath = 'C:/Users/Eric/Downloads/JMdict_e/JMdict_e.xml'
    jmdict_xml = lxml.etree.parse(jmdict_filepath)
    expr = '//keb[text()="{}"]/../..//{}'
    return jmdict_xml, expr

def extract(key, element):
    return [x.text for x in jmdict_xml.xpath(expr.format(key, element))]

def list_trim(inputlist, max_length):
    if len(inputlist) > max_length:
        return [0:max_length]
    else:
        return inputlist

for word in missing_words:
    print(list_trim(extract(word, 'reb'), 8))
    print(list_trim(extract(word, 'gloss'), 3))
    print(list_trim(extract(word, 'pos'), 3))

for word in missing_words.keys():
    missing_words[word] = [list_trim(extract(word, 'reb'), 8),
                           list_trim(extract(word, 'gloss'), 3),
                           list_trim(extract(word, 'pos'), 3),
                           date = datetime.date.today()]
	
for w in missing_words:
    find w in JMdict
    reading = JMdict value
    definition = JMdict value
    pos = JMdict value
    date = datetime.date.today()

wks.update_cell(10, 10, ['one', 'two'])
