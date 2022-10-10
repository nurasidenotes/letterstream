from __future__ import print_function
from distutils.log import error
import pathlib
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

from letterstream_app import *

# Setting Static Vars used in layout/gui
mail_types = ['First Class', 'Certified', 'Signature']

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
        sg.DropDown(values=mail_types, default_value=None, enable_events=True , key='-SET MAILTYPE-'),
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
    ## imwatchingyou.refresh_debugger()
    if event in (sg.WIN_CLOSED, 'Exit', '-CANCEL-', '-IN PROG CANCEL-', '-FINISHED-'):
        break
    # file was selected, set csv in to selected file
    if event == "-CSV IN-":
        csv_file = values["-CSV IN-"]
        continue
    if event == '-SET MAILTYPE-':
        mail_type_input = values['-SET MAILTYPE-']
        mail_type = set_mail_type(mail_type_input)
        continue
    elif event == '-RUN-':  
        window['-RUNNING-'].update(visible = True)
        window['-START-'].update(visible = False)
        window.refresh()
        if not csv_file:
            sg.popup_ok_cancel('No CSV selected. Please selected batch to run.')
            if event == 'Cancel':
                break
            if event == 'Ok':
                window['-START-'].update(visible=True)
                window['-RUNNING-'].update(visible=False)
                continue
        if not mail_type_input:
            sg.popup_ok_cancel('No mail type selected. Please select mail type before proceeding')
            if event == 'Cancel':
                break
            if event == 'Ok':
                window['-START-'].update(visible=True)
                window['-RUNNING-'].update(visible=False)
                continue
        else:
            window['-RUNNING-'].update(visible=True)
            sg.cprint('Parsing CSV data...')
            window.refresh()
            ## imwatchingyou.refresh_debugger()