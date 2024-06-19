**Companies House Data Retrieval Script**

**Overview**

This project provides a Python script to retrieve company profile, persons of significant control, and officers' details from the Companies House API for a list of company numbers provided in a CSV file. The script handles API rate limiting, retries on transient errors, and logs missing or erroneous data.

**Features**

Fetch company profile, persons of significant control, and officers' details from the Companies House API.
Handles API rate limiting (500 calls per 5 minutes).
Retries requests on server errors with exponential backoff.
Logs missing or erroneous company numbers.
Saves results into an Excel file with separate sheets for company profiles, persons of significant control, officers, and missing company numbers.
Requirements
Python 3.x
requests library
pandas library
tqdm library
openpyxl library
Installation
Install the required Python libraries:

pip install requests pandas tqdm openpyxl

Create a key.txt file in the project directory and paste your Companies House API key into this file.

Prepare a CSV file named ch.csv with a column named company_number containing the list of company numbers to query.

**Usage**

Ensure you have your key.txt (with an api key) file and ch.csv file in the project directory.
Run the script:

python main.py

Or if on windows you can use the exe which is a compiled version of the script. Simply put the ch.csv and key.txt into the same folder and double click if on windows.

**Output**

The script will generate an Excel file named company_data.xlsx in the project directory with the following sheets:

Company Profiles: Contains the company profile details.
Persons of Significant Control: Contains details of persons with significant control.
Officers: Contains the officers' details.
Missing Company Numbers: Contains company numbers that could not be retrieved along with the error codes.
Script Breakdown
RateLimiter Class
Handles rate limiting by tracking API call times and enforcing a maximum number of calls per period.

rate_limited_request Function
Makes API requests with rate limiting and retries on errors. Handles specific status codes like 429 (rate limit exceeded) and server errors (500, 502, 503, 504) with exponential backoff.

**Main Script**

Reads the list of company numbers from ch.csv.
Reads the API key from key.txt.
Iterates over each company number and makes API requests to retrieve:
Company profile
Persons of significant control
Officers' details
Logs missing or erroneous company numbers.
Saves the results into an Excel file with appropriate sheets.
Example ch.csv
csv

**company_number**

00000000
00000001
00000002
...


**License**

This project is licensed under the MIT License.

**Contributing**

Feel free to submit issues or pull requests if you have any suggestions or improvements.

**Acknowledgements**

This project uses data from the Companies House API.
