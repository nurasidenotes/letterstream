# To create and update shelf that manages company contact information to be added to letterstream.
# 1. Creates shelf if not already created.
# 3. Adds user input.
# 4. Has view to look at all current shelf information.
# 5. Has edit window to adjust company's contact.
# 6. Saves and closes shelf to be accessed by Letterstream program.

import csv

csv_file = 'resources/db/CompanyContacts.csv'

def is_company_in_csv(company):
    '''
    Finds company inside csv, returns true or false if exists
    '''
    with open(csv_file, newline='') as f:
        csvreader = csv.DictReader(f)
        for row in csvreader:
            if row['Company'] == company:
                return True
            else:
                continue
    return False

def create_return_contact(**kwargs):

    company_exists = is_company_in_csv(kwargs['company'])
    if company_exists:
        return f"Company {kwargs['company']} return info already exists. If you'd like to update the information, please press 'Update'."
    
    with open(csv_file, newline='') as f:
        contact_db = csv.DictReader(f)
        contact_fieldnames = contact_db.fieldnames
        contact_in = csv.DictWriter(f, fieldnames=contact_fieldnames)
        contact_in.writerow({
            'Company': kwargs['company'],
            'Location': kwargs['city'].join(', '+kwargs['state']),
            'Contact': kwargs['name'],
            'Phone': kwargs['phone'],
            'Email': kwargs['email'],
            'Return Contact': kwargs['return_name'],
            'Address 01': kwargs['street'],
            'Address 02': kwargs['apt'],
            'City': kwargs['city'],
            'State': kwargs['state'],
            'Zip': kwargs['zip']
        })

    return f"{kwargs['company']} created in return contact database."

def update_return_contact(**kwargs):
    # open csv, find company dict, update information with kwargs
    return True

def find_return_contact(company):
    # open csv and return matching company dict
    #errors
    contact_dict = company

    return contact_dict