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

'''
Creates functions that assign vars, format data, fills in template.
Runs PDF conversion, csv creation
Zips folder
Uploads to LetterStream via API
'''

#Static vars to be called by GUI and functions set
mail_types = ['First Class', 'Certified', 'Signature']
mt_first = 'First Class'
met_sert = 'Certified'
mt_sig = 'Signature'


template = 'LetterForwardingTemplate.docx'

def set_mail_type(input):
    '''
    Sets mail type for batch based on user input.
    input = input("Does the client want First Class, Certified, or Signature?\n")
    '''
    if input == 'First Class':
        mail_type = 'firstclass'
    elif input == 'Certified':
        mail_type = 'certnoerr'
    elif input == 'Signature':
        mail_type = 'certified'
    else:
        return 'Error: No mail type selected.'
    return mail_type

def set_csv(input):

    pass

def set_csv_out_headers():
    ouput_csv_header = {
        'UniqueDocId':'',
        'PDFFileName': '',
        'RecipientName1':'',
        'RecipientName2':'',
        'RecipientAddr1':'',
        'RecipientAddr2':'',
        'RecipientCity':'',
        'RecipientState':'',
        'RecipientZip':'',
        'SenderName1':'',
        'SenderName2':'',
        'SenderAddr1':'',
        'SenderAddr2':'',
        'SenderCity':'',
        'SenderState':'',
        'SenderZip':'',
        'PageCount':'',
        'MailType (firstclass|certified|postcard|flat)':'',
        'CoverSheet (Y|N)':'',
        'Duplex (Y|N)':'',
        'Ink (B|C)':'',
        'Paper (W(hite-default)|Y(ellow)|LB(light blue)|LG(light green)|O(range)|I(vory)|PERF(orated)':'',
        'Return Envelope (Y|N(default))':''
    }
    return ouput_csv_header

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

def create_recipient_dict(row_dict):

    recipient_dict = {
        'first_name': row_dict[' First Name'],
        'last_name': row_dict[' Last Name'],
        'address': row_dict[' Address 1'],
        'city': row_dict[' Address 1 City'],
        'state': row_dict[' Address 1 State'],
        'zip': row_dict[' Address 1 Zip'],
    }
    return recipient_dict

def create_merge_dict(company_dict, recipient_dict):
    merge_dict = {
        'FirstName':recipient_dict['first_name'],
        'LastName':recipient_dict['last_name'],
        'Address':recipient_dict['address'],
        'City':recipient_dict['city'],
        'State':recipient_dict['state'],
        'Zip':recipient_dict['zip'],
        'Company':company_dict['co_name'],
        'Co_Location':company_dict['co_location'],
        'Contact': company_dict['co_contact'],
        'Contact_Number':company_dict['co_phone'],
        'Contact_Email':company_dict['co_email']
    }
    return merge_dict