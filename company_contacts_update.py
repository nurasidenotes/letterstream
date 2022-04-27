# To create and update shelf that manages company contact information to be added to letterstream.
# 1. Creates shelf if not already created.
# 2. Opens GUI to allow user input.
# 3. Adds user input.
# 4. Has view to look at all current shelf information.
# 5. Has edit window to adjust company's contact.
# 6. Saves and closes shelf to be accessed by Letterstream program.


from __future__ import print_function
from distutils.log import error
import pathlib
from tkinter import HORIZONTAL
from tokenize import Triple
from turtle import update
from datetime import date
import time

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
import imwatchingyou