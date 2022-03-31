# 1. Take spreadsheet with names and addresses and fill in letter template.  -- CHECK
# 2. Export individual letters as PDF, naming pdf according to Name.  -- CHECK
# 3. Save PDF in project specific folder.  -- CHECK
# 4. Create lines in csv sheet to mimic API example. -- CHECK
# 5. Save csv in same folder as pdfs. -- CHECK
# 6. Zip folder. -- CHECK
# 7. Upload to letterstream/interact

from __future__ import print_function
import pathlib
from tkinter.messagebox import NO
import zipfile
from mailmerge import MailMerge
from datetime import date
import time
from docx2pdf import convert

import csv
import os
from zipfile import ZIP_DEFLATED, ZipFile
from pathlib import Path

from itertools import count
import pickle
import shelve
import requests
import base64
import hashlib
import sys

template = "single_locate.docx"
        
csv_file = 'Batch - The Ryding Company - 03-24-2022 - Appended by EmployeeLocator.csv'

# to parse csv file name into parts
file_stem = csv_file.split('.')[0]
company_name, current_date = file_stem.split(' - ')[1:3]
docx_folder_name = f"Locates_{company_name}_{current_date}_'docx'"

# To create directory (file) for word docs batch:
if not os.path.exists(docx_folder_name):
    os.mkdir(docx_folder_name)


def get_phone_types(row, index1, index2, index3):
    """
    Returns the phone types at index1, index2, index3 of the given row
    as a list of strings.
    param row: csv row of data
    param index1, index2, index3: indices of the 3 phone type columns
    returns: a list of string indicating phone types with 'C', 'R', or ''
    """
    transform = {'Mobile': 'C', 'Residential': 'R'}
    indices = [index1, index2, index3]
    return [transform.get(row[index], '') for index in indices]



# To merge csv contents with word template, create new docs, add lines to output csv
with open(csv_file, newline='') as csvfile:
    csvreader = csv.reader(csvfile)
    headers = []
    for row in csvreader: 
        if not headers:
            headers = row
            continue
        if not row[0]:
            continue
        else:
            first_name_index = headers.index(' First Name')
            last_name_index = headers.index(' Last Name')
            address_01_index = headers.index(' Address 1')
            city_01_index = headers.index(' Address 1 City')
            state_01_index = headers.index(' Address 1 State')
            zip_01_index = headers.index(' Address 1 Zip')
            address_02_index = headers.index(' Address 2')
            city_02_index = headers.index(' Address 2 City')
            state_02_index = headers.index(' Address 2 State')
            zip_02_index = headers.index(' Address 2 Zip')
            date_02_index = headers.index('Address 2 Date')
            address_03_index = headers.index(' Address 3')
            city_03_index = headers.index(' Address 3 City')
            state_03_index = headers.index(' Address 3 State')
            zip_03_index = headers.index(' Address 3 Zip')
            date_03_index = headers.index('Address 3 Date')
            address_04_index = headers.index(' Address 4')
            city_04_index = headers.index(' Address 4 City')
            state_04_index = headers.index(' Address 4 State')
            zip_04_index = headers.index(' Address 4 Zip')
            date_04_index = headers.index('Address 4 Date')	
            ptype_01_index = headers.index(' Phone 1 Type')
            phone_01_index = headers.index(' Phone 1')
            pseen_01_index = headers.index(' Phone 1 Date')
            ptype_02_index = headers.index(' Phone 2 Type')
            phone_02_index = headers.index(' Phone 2')
            pseen_02_index = headers.index(' Phone 2 date')
            ptype_03_index = headers.index(' Phone 3 Type')
            phone_03_index = headers.index(' Phone 3')
            pseen_03_index = headers.index(' Phone 3 date')
            email_01_index = headers.index(' Email 1')
            eseen_01_index = headers.index(' Email 1 date')
            email_02_index = headers.index(' Email 2')
            eseen_02_index = headers.index(' Email 2 date')
            email_03_index = headers.index(' Email 3')
            eseen_03_index = headers.index(' Email 3 date')
            phone_type_01, phone_type_02, phone_type_03 = get_phone_types(row, ptype_01_index, ptype_02_index, ptype_03_index)
            document = MailMerge(template)
            document.merge(
                firstname_01=row[first_name_index],
                lastname_01=row[last_name_index],
                address_01=row[address_01_index],
                city_01=row[city_01_index],
                state_01=row[state_01_index],
                zip_01=row[zip_01_index],
                address_02=row[address_02_index],
                city_02=row[city_02_index],
                state_02=row[state_02_index],
                zip_02=row[zip_02_index],
                lastseen_02=row[date_02_index],
                address_03=row[address_03_index],
                city_03=row[city_03_index],
                state_03=row[state_03_index],
                zip_03=row[zip_03_index],
                lastseen_03=row[date_03_index],
                address_04=row[address_04_index],
                city_04=row[city_04_index],
                state_04=row[state_04_index],
                zip_04=row[zip_04_index],
                lastseen_04=row[date_04_index],
                ptype_01 = phone_type_01,
                phone_01=row[phone_01_index],
                pseen_01=row[pseen_01_index],
                ptype_02 = phone_type_02,
                phone_02=row[phone_02_index],
                pseen_02=row[pseen_02_index],
                ptype_03 = phone_type_03,
                phone_03=row[phone_03_index],
                pseen_03=row[pseen_03_index],
                email_01=row[email_01_index],
                eseen_01=row[eseen_01_index],
                email_02=row[email_02_index],
                eseen_02=row[eseen_02_index],
                email_03=row[email_03_index],
                eseen_03=row[eseen_03_index],
            )
        document.write(f"{docx_folder_name}\Locate - {company_name} - {row[last_name_index]} - {current_date}.docx")
             # directory path only works on windows; change for mac
