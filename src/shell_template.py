import base64
import hashlib
import time

#Letterstream API Key + hash based on LetterStream API reqs FOR TESTING

api_id = 'pM73xqQl'
api_key = 'LD0cWJySLgfnsK4dfB'

def set_unique_id():
    unique_id = f'{int(time.time_ns())}'[-18:]
    return unique_id

def hash_string(unique_id, api_key):
    string_to_hash = (unique_id[-6:] + api_key + unique_id[0:6])
    encoded_string = base64.b64encode(string_to_hash.encode('ascii'))
    api_hash = hashlib.md5(encoded_string)
    hash_two = api_hash.hexdigest()
    return hash_two

def create_auth_params(hash_two, unique_id):
    auth_parameters = {
        'a': api_id,
        'h': hash_two,
        't': unique_id,
        'debug': '3'
    }
    return auth_parameters