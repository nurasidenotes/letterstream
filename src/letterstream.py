
from __future__ import print_function
from distutils.log import error
import pathlib
from platform import mac_ver, win32_ver
import shutil
from tkinter import HORIZONTAL
from tokenize import Triple
from turtle import update
import zipfile
from click import progressbar
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

import PySimpleGUI as sg
import os.path

from letterstream_app import set_company_vars, set_csv_out_headers
from shell import create_auth_params


# First the window layout
mail_type = ['First Class', 'Certified', 'Signature']

mt_first = 'First Class'
mt_cert = 'Certified'
mt_sig = 'Signature'
mail_type_input = None


file_list = [
    [
        sg.Text("Select CSV"),
        sg.In(size=(25, 1), enable_events=True, key="-CSV IN-"),
        sg.FileBrowse('Browse'),
    ],
    [
        sg.Text('Select mailtype:'),
        sg.DropDown(values=mail_type, default_value=None, enable_events=True , key='-SET MAILTYPE-'),
    ],
    [
        sg.Button('Upload to LetterStream', enable_events=True, key='-RUN-'),
        sg.Button('Cancel', enable_events=True, key='-CANCEL-'),
    ]
]
progress_window = [
    [
        sg.Text('Processing... ', key='-PROG TEXT-')
    ],
    [
        sg.Multiline('',size=(50,15), key='-PRINT-', reroute_stdout=True, reroute_stderr=True, reroute_cprint=True)
    ],
    [
        sg.Button(button_text='Cancel', enable_events=True, key='-IN PROG CANCEL-'),
        sg.Button(button_text='Finished', enable_events=True, key='-FINISHED-', visible=False),
    ],
]

# ----- Full layout -----
layout = [
        [
            sg.Column(file_list, visible=True, key='-START-'),
        ],
        [
            sg.Column(progress_window, visible=False, key='-RUNNING-'),
        ],

]
window = sg.Window('LetterStream API', layout, resizable=True)

# Run the Event Loop
while True:
    event, values = window.read()
    # Creates temp dir for running
    path = os.path.abspath(os.getcwd())
    temp_dir = os.path.join(path, 'tmp')
    if not os.path.exists(temp_dir):
        os.mkdir(temp_dir)
    if event in (sg.WIN_CLOSED, 'Exit', '-CANCEL-', '-IN PROG CANCEL-', '-FINISHED-'):
        shutil.rmtree(temp_dir)
        break
    # file was selected, set csv in to selected file
    if event == "-CSV IN-":
        csv_file = values["-CSV IN-"]
        continue
    if event == '-SET MAILTYPE-':
        mail_type_input = values['-SET MAILTYPE-']
        continue
    elif event == '-RUN-':  
        window['-RUNNING-'].update(visible = True)
        window['-START-'].update(visible = False)
        window.refresh()
        if not csv_file:
            sg.popup_ok_cancel('No CSV selected. Please select a batch to run.')
            if event == 'Cancel':
                shutil.rmtree(temp_dir)
                break
            if event == 'Ok':
                window['-START-'].update(visible=True)
                window['-RUNNING-'].update(visible=False)
                continue
        if not mail_type_input:
            sg.popup_ok_cancel('No mail type selected. Please select mail type before proceeding')
            if event == 'Cancel':
                shutil.rmtree(temp_dir)
                break
            if event == 'Ok':
                window['-START-'].update(visible=True)
                window['-RUNNING-'].update(visible=False)
                continue
        else:
            window['-RUNNING-'].update(visible=True)
            sg.cprint('Parsing CSV data...')
            window.refresh()

            dir_path = os.path.dirname(os.path.realpath(__file__))
            file_stem = csv_file.split('.')[0]
            company_name, batch_date = file_stem.split(' - ')[1:3]
            docx_folder_name = f"{company_name}_{batch_date}_'docx'"
            folder_name = f"{company_name}_{batch_date}"

            # to create copmany initials object
            initial_stem = "".join(item[0].upper() for item in company_name.split())

            # To create variable letter template from Company information
            contact_csv = os.path.join(dir_path, 'CompanyContacts.csv')

            # Assigns template
            template = os.path.join(dir_path, "LetterForwardingTemplate.docx")

            # Creates temp dir for running
            temp_dir = os.path.join(path, 'tmp')
            if not os.path.exists(temp_dir):
                os.mkdir(temp_dir)

            #output_path = os.path.dirname()
            sg.cprint('Creating Batch Folders...')
            window.refresh()
            # To create directory (file) for word docs batch:
            docx_directory = os.path.join(temp_dir, docx_folder_name)
            if not os.path.exists(docx_directory):
                os.mkdir(docx_directory)

            # To create directory (file) for full batch:
            folder = os.path.join(temp_dir, folder_name)
            if not os.path.exists(folder):
                os.mkdir(folder)

            sg.cprint('Consolidating previous orders...')
            window.refresh()
            #creates variable counter that saves with co as key to shelf "unique_id"
            id_count = ''
            s_key = f"{initial_stem}"
            order_num = ''
            s_order_key = f'{initial_stem}_order'
            order_count = 1

            s = shelve.open("unique_id")

            if not s.get(s_key):
                id_count = 1
            elif s.get(s_key) is not None:
                id_count = s.get(s_key)

            if not s.get(s_order_key):
                order_num = 1
            elif s.get(s_order_key) is not None:
                order_num = s.get(s_order_key)+1

            s.close()

            # # creates output header list
            # output_csv_header = []
            sg.cprint('Assigning static variables...')
            window.refresh()
            output_csv_header = set_csv_out_headers()
            # opens output template and pulls headers
            # with open('Batch_Template.csv', newline='') as output_csv:
            #     output_reader = csv.DictReader(output_csv)
            #     output_csv_header = output_reader.fieldnames
                    
            # creates output csv in output folder
            output_csv = os.path.join(folder, f'{company_name}_Batch{order_num:0>3}_{batch_date}.csv')

            # creates batch dictionary for csv output
            batch_column_dict = dict.fromkeys(output_csv_header , None)

            #function that takes company input and creates a dict 
            def set_company_vars(company_csv):
                company_dict = {
                    'co_name': company_csv['Company'],
                    'co_location' : company_csv['Location'],
                    'co_contact' : company_csv['Contact'],
                    'co_phone' : company_csv['Phone'],
                    'co_email' : company_csv['Email'],
                    'co_return' : company_csv['Return Contact'],
                    'co_address_01' : company_csv['Address 01'],
                    'co_address_02' : company_csv['Address 02'],
                    'co_city' : company_csv['City'],
                    'co_state' : company_csv['State'],
                    'co_zip' : company_csv['Zip']
                }
                return company_dict

            # set company variables based on contact_csv
            with open(contact_csv, newline='') as companycsv:
                csvreader = csv.DictReader(companycsv)
                for row in csvreader: 
                    if row['Company'] != company_name:
                        continue
                    elif row['Company'] == company_name:
                        company_dict = set_company_vars(row)
                        break

            
            sg.cprint('Setting Mail Type')
            window.refresh()
            # sets mail type for batch
            ## mail_type_input = input("Does the client want First Class, Certified, or Signature?\n")
            if mail_type_input == 'First Class':
                mail_type_input = 'firstclass'
            elif mail_type_input == 'Certified':
                mail_type_input = 'certnoerr'
            elif mail_type_input == 'Signature':
                mail_type_input = 'certified'

            # sets static company/batch variables to output csv dict
            batch_column_dict.update({
                'SenderName1': f'{company_name}',
                'SenderName2': 'Attn: {}'.format(company_dict['co_contact']),
                'SenderAddr1': company_dict['co_address_01'],
                'SenderAddr2': company_dict['co_address_02'],
                'SenderCity': company_dict['co_city'], 
                'SenderState': company_dict['co_state'], 
                'SenderZip': company_dict['co_zip'], 
                'PageCount': '1', 
                'MailType (firstclass|certified|postcard|flat)': f'{mail_type_input}', 
                'CoverSheet (Y|N)': 'Y', 
                'Duplex (Y|N)': 'N', 
                'Ink (B|C)': 'B', 
                'Paper (W(hite-default)|Y(ellow)|LB(light blue)|LG(light green)|O(range)|I(vory)|PERF(orated)': 'W', 
                'Return Envelope (Y|N(default))': 'N'
            })
            
            # create function that checks zip code for correct # of zeroes
            def check_zip(zip):
                new_zip = ''
                if len(zip) < 5:
                    if len(zip) == 4:
                        new_zip = f'0{zip}'
                        return new_zip
                    elif len(zip) == 3:
                        new_zip = f'00{zip}'
                        return new_zip
                    elif len(zip) == 2:
                        new_zip = f'000{zip}'
                        return new_zip
                    elif len(zip) == 1:
                        new_zip = f'0000{zip}'
                        return new_zip
                    else:
                        return error('Zip code error.')
                else:
                    return f'{zip}'

            # create function that updates output csv dict with variables in with
            def generate_batch_row():
                batch_column_dict.update({
                    'UniqueDocId': f'{doc_id}',
                    'PDFFileName': f'{pdf_name}',
                    'RecipientName1': f'{row[first_name_index]} {row[last_name_index]}',
                    'RecipientAddr1': f'{row[address_index]}',
                    'RecipientAddr2': None,
                    'RecipientCity': f'{row[city_index]}',
                    'RecipientState': f'{row[state_index]}',
                    'RecipientZip': zip_code
                })
                return

            with open(csv_file, newline='') as csvnumber:
                csv_number = csv.reader(csvnumber)
                data = [row for row in csv_number]
                total_rows = len(data)

            sg.cprint('Parsing data and creating documents:')
            window.refresh()
            # To merge csv contents with word template, create new docs, add lines to output csv
            with open(csv_file, newline='') as csvfile, open(output_csv, 'w', encoding='UTF8', newline='') as batch_output:
                csvreader = csv.reader(csvfile)
                csvwriter = csv.DictWriter(batch_output, fieldnames=output_csv_header)
                csvwriter.writeheader()
                headers = []
                doc_count = 1
                row_is = 1
                for row in csvreader: 
                    row_is += 1
                    sg.OneLineProgressMeter('Writing: ', row_is, total_rows)
                    window.refresh()
                    if not headers:
                        headers = row
                        continue
                    if not row[0]:
                        continue
                    if not row[headers.index('Address 1')]:
                        continue
                    else:
                        first_name_index = headers.index('First Name')
                        last_name_index = headers.index('Last Name')
                        address_index = headers.index('Address 1')
                        city_index = headers.index('Address 1 City')
                        state_index = headers.index('Address 1 State')
                        zip_index = headers.index('Address 1 Zip')
                        nospace = row[last_name_index]
                        doc_id = f"{batch_date}_{initial_stem}{id_count:0>4}"
                        pdf_name = f'{order_count:0>4}_{nospace.replace(" ","")}_{row[first_name_index]}.pdf'
                        doc_name = f'{order_count:0>4}_{nospace.replace(" ","")}_{row[first_name_index]}.docx'
                        zip_check = row[zip_index]
                        zip_code = check_zip(zip_check)
                        sg.cprint(f'Creating {doc_name}')
                        window.refresh()
                        zip_check = row[zip_index]
                        zip_code = check_zip(zip_check)
                        generate_batch_row()
                        csvwriter.writerow(batch_column_dict.copy())
                        document = MailMerge(template)
                        document.merge(
                            FirstName=row[first_name_index],
                            LastName=row[last_name_index],
                            Address=row[address_index],
                            City=row[city_index],
                            State=row[state_index],
                            Zip=zip_code,
                            Company=company_dict['co_name'],
                            Co_Location=company_dict['co_location'],
                            Contact=company_dict['co_contact'],
                            Contact_Number=company_dict['co_phone'],
                            Contact_Email=company_dict['co_email']
                        )
                        id_count += 1
                    document.write(f"{docx_directory}/{doc_name}")
                    ##MAC document.write(f"{docx_folder_name}/{doc_name}")
                    order_count += 1
                
                # docx2pdf convert folder of docx to other folder of pdf
                sg.cprint('Converting to PDF:')
                window.refresh()
                convert(docx_directory,folder)
                window.refresh()
                sg.cprint('PDF conversion complete.')
                window.refresh()


            # outputs last id_count number to dict that saves to shelf
            current_file_count = {s_key:id_count}
            current_order_num = {s_order_key:order_num}
            s=shelve.open("unique_id")
            s.update(current_file_count)
            s.update(current_order_num)
            s.close()

            ## hide when executable?
            #clear_shelf = sg.popup_yes_no(f"Would you like to clear {company_name}'s document count?")
            #clear_shelf_order = sg.popup_yes_no(f"Would you like to clear {company_name}'s order count?")

            # #  clears current company's letterstream id_count, if requested
            ## clear_shelf = input(f'Clear {company_name} LetterStream count? y/n: ')
            # if clear_shelf == 'Yes':
            #     s=shelve.open("unique_id")
            #     s.pop(s_key)
            #     s.close()
            # else:
            #     pass

            # #  clears current company's letterstream id_count, if requested
            ##clear_shelf_order = input(f'Clear {company_name} order count? y/n: ')
            # if clear_shelf_order == 'Yes':
            #     s=shelve.open("unique_id")
            #     s.pop(s_order_key)
            #     s.close()
            # else:
            #     pass
            
            sg.cprint('Zipping folder for upload...')
            window.refresh()

            # zips folder
            zip_file = f'{folder_name}.zip'
            zip_directory = pathlib.Path(folder)

            with zipfile.ZipFile(zip_file, 'w', ZIP_DEFLATED, allowZip64=True) as z:
                for f in zip_directory.iterdir():
                    z.write(f, arcname=f.name,)
            
            sg.popup_ok('Files ready for upload.')
            window.refresh()

            sg.cprint('Uploading to Letterstream:')
            window.refresh()

            # Authenticates letterstream api connection using random variables
            sg.cprint('Authenticating user...')
            window.refresh()

            auth_parameters = create_auth_params()
            
            sg.cprint('Sending .zip to Letterstream...')
            window.refresh()
            with open(zip_file, 'rb') as fileobj:
                r = requests.post(url='https://www.letterstream.com/apis/index.php',data=auth_parameters, files={'multi_file': (zip_file, fileobj)})
                print(r.status_code)
                if "AUTHOK" in r.text:
                    print(r.text)
                    window.refresh()
                    sg.popup_ok('Batch upload successful.')
                    window['-FINISHED-'].update(visible=True)
                elif not ("AUTHOK" in r.text):
                    sg.popup_error('ERROR during upload')
                    print(f"ERROR on {zip_file}: ")
                    print(r.text)

## parse r.text to be legible -- json?
## executable package