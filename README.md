# BBO myhands parser

Scripts for parsing BBO myhands data - made out of the frustration of manually copying data into sheets.

## What does it do?

For now, it transforms myhands html to an xlsx file, formatted for Kit's method of cheating detection described in the article
[Detecting Illicit Knowledge](http://bridgewinners.com/article/view/determining-illicit-knowledge/)

## Prereq's
You'll need python3 and some modules which are easily installed using pip3 (anyone can do this, no knowledge of Python required). 

## Usage

Clone the project, download myhands page into html folder under the default name (Bridge Base Online - Myhands.html)
```bash
python3 html2xlsx_kit.py
```
## Twisting the dates
Myhands allows limited temporal querying options. If you want data from particular periods (say 2 weeks), you can edit the timestamp in the address bar - there are online tools for timestamp such as [this](https://www.timestampconvert.com/). Also (and especially if the data is overlapped), you should edit html manually (just throw out the rows from the tournaments you don't need).
