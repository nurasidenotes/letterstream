# 1. Take spreadsheet with names and addresses and fill in letter template.  -- CHECK
# 2. Export individual letters as PDF, naming pdf according to Name.  -- CHECK
# 3. Save PDF in project specific folder.  -- CHECK
# 4. Create lines in csv sheet to mimic API example. -- CHECK
# 5. Save csv in same folder as pdfs. -- CHECK
# 6. Zip folder.
# 7. Upload to letterstream/interact

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from docx2pdf import convert

import csv
import os
from pathlib import Path

from itertools import count
import pickle
import shelve

template = "LetterForwardingTemplate.docx"
        
csv_file = 'Batch - Kalamazoo Country Club - 03-09-2022 - Appended by EmployeeLocator.csv'

# to parse csv file name into parts
file_stem = csv_file.split('.')[0]
company_name, current_date = file_stem.split(' - ')[1:3]
docx_folder_name = f"{company_name}_{current_date}_'docx'"
folder_name = f"{company_name}_{current_date}"

# to create copmany initials object
initial_stem = "".join(item[0].upper() for item in company_name.split())

# To create variable letter template from Company information
contact_csv = 'CompanyContacts.csv'

# To create directory (file) for word docs batch:
if not os.path.exists(docx_folder_name):
    os.mkdir(docx_folder_name)

# To create directory (file) for full batch:
if not os.path.exists(folder_name):
    os.mkdir(folder_name)

# creates output header list
output_csv_header = []

# opens output template and pulls headers
with open('Batch_Template.csv', newline='') as output_csv:
   output_reader = csv.DictReader(output_csv)
   output_csv_header = output_reader.fieldnames
        
# creates output csv variable object
output_csv = f"{folder_name}\LetterStream_Batch_{company_name}_{current_date}.csv"


# creates batch dictionary for csv output
batch_column_dict = dict.fromkeys(output_csv_header , None)


# company variables 
# 'Company', 'Location', 'Contact', 'Phone', 'Email', 'Return Contact', 'Address 01', 'Address 02', 'City', 'State', 'Zip'
co_name = ''
co_location = ''
co_contact = ''
co_phone = ''
co_email = ''
co_return = ''
co_address_01 = ''
co_address_02 = ''
co_city = ''
co_state = ''
co_zip = ''

# set company variables based on contact_csv
with open(contact_csv, newline='') as csvfile:
    csvreader = csv.reader(csvfile)
    headers = []
    for row in csvreader: 
        if not headers:
            headers = row
            continue
        if row[headers.index('Company')] == company_name:
            co_name = row[headers.index('Company')]
            co_location = row[headers.index('Location')]
            co_contact = row[headers.index('Contact')]
            co_phone = row[headers.index('Phone')]
            co_email = row[headers.index('Email')]
            co_return = row[headers.index('Return Contact')]
            co_address_01 = row[headers.index('Address 01')]
            co_address_02 = row[headers.index('Address 02')]
            co_city = row[headers.index('City')]
            co_state = row[headers.index('State')]
            co_zip = row[headers.index('Zip')]
            break

# sets mail type for batch
mail_type_input = input("Does the client want First Class, Certified, or Signature?\n")
if mail_type_input == 'First Class':
    mail_type_input = 'firstclass'
elif mail_type_input == 'Certified':
    mail_type_input = 'certnoerr'
elif mail_type_input == 'Signature':
    mail_type_input = 'certified'

# sets static company/batch variables to output csv dict
batch_column_dict.update({
    'SenderName1': f'{company_name}',
    'SenderName2': f'Attn: {co_return}',
    'SenderAddr1': f'{co_address_01}',
    'SenderAddr2': f'{co_address_02}',
    'SenderCity': f'{co_city}', 
    'SenderState': f'{co_state}', 
    'SenderZip': f'{co_zip}', 
    'PageCount': '1', 
    'MailType (firstclass|certified|postcard|flat)': f'{mail_type_input}', 
    'CoverSheet (Y|N)': 'Y', 
    'Duplex (Y|N)': 'N', 
    'Ink (B|C)': 'B', 
    'Paper (W(hite-default)|Y(ellow)|LB(light blue)|LG(light green)|O(range)|I(vory)|PERF(orated)': 'W', 
    'Return Envelope (Y|N(default))': 'N'
})

#creates variable counter that saves with co as key to shelf "unique_id"
id_count = ''
s_key = f"{initial_stem}"

s = shelve.open("unique_id")

if not s.get(s_key):
    id_count = 1
elif s.get(s_key) is not None:
    id_count = s.get(s_key)

s.close()

# create function that updates output csv dict with variables in with
def generate_batch_row():
    batch_column_dict.update({
        'UniqueDocId': f'{doc_id}',
        'PDFFileName': f'{pdf_name}',
        'RecipientName1': f'{row[first_name_index]}',
        'RecipientName2': f'{row[last_name_index]}',
        'RecipientAddr1': f'{row[address_index]}',
        'RecipientAddr2': None,
        'RecipientCity': f'{row[city_index]}',
        'RecipientState': f'{row[state_index]}',
        'RecipientZip': f'{row[zip_index]}'
    })
    return

# To merge csv contents with word template, create new docs, add lines to output csv
with open(csv_file, newline='') as csvfile, open(output_csv, 'w', encoding='UTF8', newline='') as batch_output:
    csvreader = csv.reader(csvfile)
    csvwriter = csv.DictWriter(batch_output, fieldnames=output_csv_header)
    csvwriter.writeheader()
    headers = []
    for row in csvreader: 
        if not headers:
            headers = row
            continue
        if not row[0]:
            continue
        if not row[headers.index(' Address 1')]:
            continue
        else:
            first_name_index = headers.index(' First Name')
            last_name_index = headers.index(' Last Name')
            address_index = headers.index(' Address 1')
            city_index = headers.index(' Address 1 City')
            state_index = headers.index(' Address 1 State')
            zip_index = headers.index(' Address 1 Zip')
            doc_id = f"{current_date}_{initial_stem}_{id_count:0>4}"
            pdf_name = f'{company_name}_{row[last_name_index]}, {row[first_name_index]}_{current_date}.pdf'
            generate_batch_row()
            csvwriter.writerow(batch_column_dict.copy())
            document = MailMerge(template)
            document.merge(
                FirstName=row[first_name_index],
                LastName=row[last_name_index],
                Address=row[address_index],
                City=row[city_index],
                State=row[state_index],
                Zip=row[zip_index],
                Company=co_name,
                Co_Location=co_location,
                Contact=co_contact,
                Contact_Number=co_phone,
                Contact_Email=co_email
            )
            id_count += 1
        document.write(f"{docx_folder_name}\{company_name}_{row[last_name_index]}, {row[first_name_index]}_{current_date}.docx")
             # directory path only works on windows; change for mac
    
    # docx2pdf convert folder of docx to other folder of pdf
    convert(f"{docx_folder_name}/",f"{folder_name}/")

# outputs last id_count number to dict that saves to shelf
current_file_count = {f"{initial_stem}":id_count}
s=shelve.open("unique_id")
s.update(current_file_count)
s.close()

<<<<<<< Updated upstream
<<<<<<< Updated upstream
# #  clears current company's letterstream id_count
# s=shelve.open("unique_id")
# s.pop(s_key)
# s.close()

# ToDo Zip folder
# ToDo Upload zipped folder to LetterStream API/server 
# ToDo Create loop to run program and go through set number of csv files?\
#       \program that will loop through files runs this one?
# ToDo Add variable to the csv file input -- select file at beginning of program?



# dict to associate field_name_in_merge_template : column_name_in_csv
# merge_field_lookup = {
#     'City': ' Address 1 City', 
#     'Address': ' Address 1', 
#     'Zip': ' Address 1 Zip', 
#     'LastName': ' Last Name', 
#     'State': ' Address 1 State', 
#     'FirstName': ' First Name'
# }
# # Defines function "generate_batch_row" to grab recipient list and fill in csv
# def generate_batch_row(recipient):
#     """"
#                 FirstName=row[first_name_index],
#                 LastName=row[last_name_index],
#                 Address=row[address_index],
#                 City=row[city_index],
#                 State=row[state_index],
#                 Zip=row[zip_index],
#                 Company=co_name,
#                 Co_Location=co_location,
#                 Contact=co_contact,
#                 Contact_Number=co_phone,
#                 Contact_Email=co_email
#                 PDFFileName='' 
#     """
#     batch_column_dict = {
#         "UniqueDocId": f"{current_date}_{initial_stem}_{id_count:0>4}",
#         "PDFFileName": "PDFFileName",
#         "RecipientName1": None,
#         "RecipientName2": None,
#         "RecipientAddr1": None,
#         "RecipientAddr2": None,
#         "RecipientCity": None,
#         "RecipientState": None,
#         "RecipientZip": None,
#         "SenderName1": None,
#         "SenderName2": None,
#         "SenderAddr1": None,
#         "SenderAddr2": None,
#         "SenderCity": None,
#         "SenderState": None,
#         "SenderZip": None,
#         "PageCount": "1",
#         "MailType (firstclass|certified|postcard|flat)": None,
#         "CoverSheet (Y|N)": "Y",
#         "Duplex (Y|N)": "N",
#         "Ink (B|C)": "B",
#         "Paper (W(hite-default)|Y(ellow)|LB(light blue)|LG(light green)|O(range)|I(vory)|PERF(orated)": "W",
#         "Return Envelope (Y|N(default))": "N"
#     }
   



# ToDo Create loop to run program and go through set number of csv files?\
#       \program that will loop through files runs this one?
=======
=======
>>>>>>> Stashed changes
clear_shelf = input(f'Clear {company_name} LetterStream count? y/n: ')
if clear_shelf == 'y':
    s=shelve.open("unique_id")
    s.pop(s_key)
    s.close()
else:
    pass

# #  clears current company's letterstream id_count, if requested
clear_shelf_order = input(f'Clear {company_name} order count? y/n: ')
if clear_shelf_order == 'y':
    s=shelve.open("unique_id")
    s.pop(s_order_key)
    s.close()
else:
    pass

# zips folder
zip_file = f'{folder_name}.zip'
zip_directory = pathlib.Path(f'{folder_name}/')

with zipfile.ZipFile(zip_file, 'w', ZIP_DEFLATED, allowZip64=True) as z:
    for f in zip_directory.iterdir():
        z.write(f, arcname=f.name,)


# LetterStream API id/key:
## Your API_ID : dN26vwWd
## Your API_KEY : TP6bKLpVFgqcrL2wrM

# Authenticates letterstream api connection using random variables

api_id = 'dN26vwWd'
api_key = 'TP6bKLpVFgqcrL2wrM'
unique_id = f'{int(time.time_ns())}'[-18:]
string_to_hash = (unique_id[-6:] + api_key + unique_id[0:6])

encoded_string = base64.b64encode(string_to_hash.encode('ascii'))
api_hash = hashlib.md5(encoded_string)
hash_two = api_hash.hexdigest()

auth_parameters = {
    'a': api_id,
    'h': hash_two,
    't': unique_id,
    'debug': '3'
}

with open(zip_file, 'rb') as fileobj:
    r = requests.post(url='https://www.letterstream.com/apis/index.php',data=auth_parameters, files={'multi_file': (zip_file, fileobj)})
    print(r.status_code)
    print(r.text)

>>>>>>> Stashed changes
