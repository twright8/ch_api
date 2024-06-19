import requests
import pandas as pd
import time
from tqdm import tqdm
from openpyxl import Workbook
from collections import deque
from datetime import datetime, timedelta


# Rate limiter
class RateLimiter:
    def __init__(self, max_calls, period):
        self.max_calls = max_calls
        self.period = timedelta(seconds=period)
        self.call_times = deque()

    def acquire(self):
        current_time = datetime.now()
        while self.call_times and (current_time - self.call_times[0] > self.period):
            self.call_times.popleft()
        if len(self.call_times) < self.max_calls:
            self.call_times.append(current_time)
        else:
            wait_time = (self.period - (current_time - self.call_times[0])).total_seconds()
            time.sleep(wait_time)
            self.acquire()


# Instantiate the rate limiter
rate_limiter = RateLimiter(500, 300)


def rate_limited_request(api_url, api_key, max_retries=5):
    rate_limiter.acquire()
    for i in range(max_retries):
        response = requests.get(api_url, auth=(api_key,''))
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 429:
            wait_time = int(response.headers.get('Retry-After', 2 ** i))
            print(f"Hit rate limit, pausing for {wait_time:.2f}")
            time.sleep(wait_time)
        elif response.status_code in {500, 502, 503, 504}:
            # Server errors, retry after backoff
            wait_time = 2 ** i
            print(f"Rate limit exceeded. Waiting for {wait_time:.2f} seconds.")
            time.sleep(wait_time)
        else:
            # For other errors, log the status code and return None
            return {'status_code': response.status_code}
    return None



# Read the company numbers from the CSV file
company_numbers = pd.read_csv('ch.csv')['company_number'].tolist()

# Read the API key from the text file
with open('key.txt', 'r') as file:
    api_key = file.read().strip()


base_url = "https://api.company-information.service.gov.uk/company/"

company_profiles = []
persons_significant_control = []
officers = []
missing_numbers = []

for company_number in tqdm(company_numbers, desc="Processing companies"):
    profile_url = f"{base_url}{company_number}"
    psc_url = f"{base_url}{company_number}/persons-with-significant-control"
    officers_url = f"{base_url}{company_number}/officers"

    profile_data = rate_limited_request(profile_url, api_key)
    psc_data = rate_limited_request(psc_url, api_key)
    officers_data = rate_limited_request(officers_url, api_key)

    if profile_data and 'status_code' not in profile_data:
        company_profiles.append(profile_data)
    else:
        missing_numbers.append((company_number, profile_data.get('status_code', 'Unknown')))

    if psc_data and 'status_code' not in psc_data:
        persons_significant_control.append(psc_data)

    if officers_data and 'status_code' not in officers_data:
        officers.append(officers_data)

# Save results to Excel file
with pd.ExcelWriter('company_data.xlsx') as writer:
    pd.json_normalize(company_profiles).to_excel(writer, sheet_name='Company Profiles')
    pd.json_normalize(persons_significant_control).to_excel(writer, sheet_name='Persons of Significant Control')
    pd.json_normalize(officers).to_excel(writer, sheet_name='Officers')
    pd.DataFrame(missing_numbers, columns=['Company Number', 'Error Code']).to_excel(writer,
                                                                                     sheet_name='Missing Company Numbers')


