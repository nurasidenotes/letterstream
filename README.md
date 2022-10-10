# letterstream_app

Automates letter forwarding process using docx mail-merge, docxtopdf, and API integration with LetterStream.com

Letterstream: Opens a GUI with a file browser and a drop down menu to choose mailing type. Matches company name in file name to CompanyContacts.csv, assigning the company return address, contact, and template dictionaries. Pulls recipient data from selected csv, checks to see if recipient has address. If there is an eligible address, information is parsed into a dictionary, formatted, and added to a docx template in a batch specific folder. The documents created are then converted to pdf, added to a new batch folder, the required information added to a new csv, and the folder with the pdfs and csv zipped to be uploaded to LetterStream via their API. A pop up confirms that you are ready to upload the files. The program then authenticates the EL account with LetterStream and posts the zipped folder to the server. The API response is printed to the window so that the user can verify that the batch was posted successfully before closing the program.

Public test version; executable package has sample 'authentication' for API and runs through mock API enpoints, returning the same messages.

## Directory Structure
```
README.md
requirements.txt
src/
 letterstream.py
 Batch_Template.csv
 LetterForwardingTemplate.docx
 CompanyContacts.csv
test/
 test_letterstream.py
datasets/
 Batch - Company - date - Appended by EmployeeLocator.csv
 LetterstreamAPI.pdf
```
All source code is in the `src/` directory. Tests are in the `test/` directory. There is one test file per source code file, and the test file name format is `<source file name>_test`.py.

## Running Executable
This project is available to download as an executable package.

1. Download 'LetterstreamApp' package
2. Download datasets/* from directory 
3. Open 

## Getting started for local development
This project is written with python3. 
1. Create and activate a virtual environment.
```
python3 -m venv .venv
```
2. Activate virtual environment.
On Linux and MacOS:
```
source .venv/bin/activate
```
On Windows:
```
.venv\Scripts\activate.bat
```
3. Install project dependencies.
```
pip3 install requirements.txt
```
4. Create shell script using template.

5. Make sure there is working Microsoft Word installed.

5. Run scripts.
```
python3 src/letterstream.py
```
6. Input Batch - Company - Date - Appended by EmployeeLocator.csv from /datasets.
7. Press 'Upload Batch'
8. Follow prompts.

## Running tests
To run all tests, activate the virtual environment and run the following from the repo root directory.
```
python3 -m pytest
```
To run a specific test file, pass the path to the test file as an argument:
```
python3 -m pytest test/test_letterstream.py
```
