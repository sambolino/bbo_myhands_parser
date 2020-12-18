#!/usr/bin/env python3

"""
A script for conversion of BBO myhands data to an xlsx file
"""

import os
from bs4 import BeautifulSoup as bs
import urllib.request as urlreq
import re
import xlsxwriter


abspath = os.path.abspath(os.getcwd())

#open manually downloaded html
html_page = urlreq.urlopen("file://%s/html/Bridge Base Online - Myhands.html" % (abspath))
soup = bs(html_page, features="html.parser")
html_page.close()
#searching for movies, dates and tournaments (if provided)
tags = soup.findAll(
        lambda tag:tag.name == "th" and 'colspan' in tag.attrs and bool(re.search('^[\d\-]+$', tag.text))==True
            or (tag.name =="a" and tag.text == 'Movie')
            or (tag.name=='td' and tag.has_attr('colspan') and tag.attrs['colspan'] == '6'))

rows = []
date = tournament = movie = ''

#generate list of dictionaries
for tag in tags:
    if bool(re.search('^[\d\-]+$', tag.text))==True:
        date=tag.get_text()
    elif tag.name=='td' and tag.has_attr('colspan') and tag.attrs['colspan'] == '6':
        tournament = tag.findChild().get("href")
    else:
        movie = tag.get("href")
        row = {'date':date, 'tournament':tournament, 'movie':movie}
        rows.append(row)

#now populate and format xlsx
workbook = xlsxwriter.Workbook('xlsx/myhands.xlsx')
worksheet = workbook.add_worksheet()

irow = 0
icol = 0

format = workbook.add_format()

format.set_bold()
format.set_pattern(1)
format.set_color('white')
format.set_bg_color('brown')
worksheet.write_row("A1:C1", ['Date','Tournament','Movie'], format)
worksheet.write_row("D1:F1", ['Incriminating','Absolving','Neutral'], format)
worksheet.write_row("G1:H1", ['Comment1','Comment2'], format)
worksheet.set_column(0, 1, 12)
worksheet.set_column(2, 2, 88)
worksheet.set_column(3, 5, 13)
worksheet.set_column(6, 7, 30)

for row in rows:
    irow += 1
    worksheet.write(irow, icol, row['date'])
    worksheet.write(irow, icol + 1, row['tournament'])
    worksheet.write(irow, icol + 2, row['movie'])

irow +=4
worksheet.write(irow, 3, "=SUM(D2:D%s)" % (irow-3,))
worksheet.write(irow, 4, "=SUM(E2:E%s)" % (irow-3,))
worksheet.write(irow, 5, "=SUM(F2:F%s)" % (irow-3,))
irow +=1
worksheet.write(irow, 4, "incriminating %", format)
worksheet.write(irow, 5, "=Round(D%s/SUM(D%s, E%s), 2)" % (3*(irow,)))
irow +=1
worksheet.write(irow, 4, "meaningful %", format)
worksheet.write(irow, 5, "=Round(SUM(D%s, E%s)/SUM(D%s, E%s, F%s), 2)" % (5*(irow-1,)))

workbook.close()
