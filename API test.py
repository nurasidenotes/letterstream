import requests
import base64
import hashlib
import datetime
import time
import http.client
import json

# LetterStream API id/key:
## Your API_ID : dN26vwWd
## Your API_KEY : TP6bKLpVFgqcrL2wrM

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

auth = requests.post(url='https://www.letterstream.com/apis/index.php',data=auth_parameters)

print(auth.text)


# encoded_string = base64.b64encode(string_to_hash.encode('ascii'))
# hash = hashlib.md5(encoded_string)

# auth_parameters = {
#     'a': api_id,
#     'h': hash,
#     't': unique_id,
#     'debug': '3'
# }

# ls_api = http.client.HTTPSConnection('www.letterstream.com')

# ls_api.request("POST", '/apis/index.php', json.dumps(auth_parameters))
# response = ls_api.getresponse()
# result = response.read()

# print(result.decode("utf-8"))

##auth = requests.request('POST', 'https://www.letterstream.com/apis/index.php', params=auth_parameters)
##stat = auth.status_code
##print(auth.text)

##auth = requests.request('GET', 'https://www.letterstream.com/apis/index.php', params=auth_parameters)

##print(response.text)
## https://www.letterstream.com/apis/index.php?
## https://www.letterstream.com/apis/index.php?a=$api_id&h=$hash&t=$unique_id&debug=3



## sample:
# r = requests.get('https://ws.audioscrobbler.com/2.0/', headers=headers, params=payload)
#r.status_code