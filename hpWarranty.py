import pandas as pd
import xlrd
import xlwt
import requests
import simplejson as json
import time
import dateutil.parser
import datetime

# contains apiKey and apiSecret, will have to generate your own from
# https://developers.hp.com/hp-product-warranty/api/product-warranty-api
import config

def _url(path):
    """
    HP Warranty API's base url
    """
    return 'https://css.api.hp.com' + path

def import_xlsx(xlsx_filename):
    """
    Requires an .xlsx file with a column titled Serial

    Isolates the Serial column from the .xlsx file and returns it in a pandas dataframe
    """
    df = pd.read_excel(xlsx_filename)
    data = []
    for row in df.Serial.iteritems():
        serial = { 'sn': row[1] }
        data.append(serial)
    return data

def get_access_token():
    """
    Requires config.py to have a valid apiKey and apiSecret

    Sends a request for and returns an access token
    """
    tokenBody = {
        'apiKey': config.apiKey,
        'apiSecret': config.apiSecret,
        'grantType': 'client_credentials',
        'scope': 'warranty'
    }
    tokenHeaders = { 'Accept': 'application/json' }
    tokenResponse = requests.post(
        _url('/oauth/v1/token'), 
        data = tokenBody, 
        headers = tokenHeaders
    )
    tokenJson = tokenResponse.json()
    return tokenJson['access_token']

input_spreadsheet = 'test.xlsx'
print('=== Importing Excel Spreadsheet ===')
serial_numbers = import_xlsx(input_spreadsheet)
print('=== Imported, subset of data ===')
print(serial_numbers[:2])
print('=== Getting Access Token ===')
token = get_access_token()
jobHeaders = {
        'Accept': 'application/json',
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
}

print('Creating new batch job...')
jobResponse = requests.post(
    _url('/productWarranty/v2/jobs/'), 
    data=json.dumps(serial_numbers), 
    headers=jobHeaders
)
job = jobResponse.json()
print('Batch job created successfully.')
print('Job ID: ' + job['jobId'])
print('Estimated time in seconds to completion: ' + str(job['estimatedTime']))
print('')

if (job['estimatedTime'] > 1200):
    time.sleep(40)
else:
    time.sleep(20)
headers = {
    'Authorization': 'Bearer ' + token,
    'Accept-Encoding': 'gzip,deflate'
}
status = 'incomplete'
while (status == 'incomplete'):
    monitorResponse = requests.get(_url('/productWarranty/v2/jobs/' + job['jobId']), headers=headers)
    monitor = monitorResponse.json()
    if (monitor['status'] != "completed"):
        if (monitor['estimatedTime'] > 1200):
            print('Estimated time in seconds to completion: ' + str(monitor['estimatedTime']) + '\nNext job check in 10 minutes...\n')
            time.sleep(200)
        elif (monitor['estimatedTime'] > 600):
            print('Estimated time in seconds to completion: ' + str(monitor['estimatedTime']) + '\nNext job check in 5 minutes...\n')
            time.sleep(100)
        else:
            print('Estimated time in seconds to completion: ' + str(monitor['estimatedTime']) + '\nNext job check in 1 minute...\n')
            time.sleep(10)
    else:
        status = 'complete'

resultsResponse = requests.get(_url('/productWarranty/v2/jobs/' + job['jobId'] + '/results'), headers=headers)
results = resultsResponse.json()
today = datetime.date.today()

for r in results:
    serialNumber = r["product"]["serialNumber"]
    print("\n=====New Product=====")
    for offer in r["offers"]:
        if "HP HW Maintenance Onsite Support" in offer["offerDescription"]:
            print("Warranty Started: " + offer["serviceObligationLineItemStartDate"])
            print("Warranty Ended: " + offer["serviceObligationLineItemEndDate"])
            parsed = dateutil.parser.parse(offer['serviceObligationLineItemEndDate']).date()
            if today < parsed:
                print('Warranty active for ' + serialNumber + "\n")
            else:
                print('Warranty inactive for ' + serialNumber + "\n")

try:
    f = open(job['jobId'] + '.json', 'w')
    f.write(json.dumps(results))
    print("\nWarranty information was retrieved for " + str(len(results)) + " objects.\nTo view raw data, see " + job['jobId'] + ".json.")
    f.close()
except Exception:
    print("\nWarranty information was retrieved for " + str(len(results)) + ' objects.\nRaw data could not be written to file.')