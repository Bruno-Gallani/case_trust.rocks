The solution was all written in **Python**, the original file used for solving the problem was a jupyter notebook `.ipynb`. You can find it <a href="./Case - TrustRocks.ipynb">here</a>.
- The official <a href="https://developers.google.com/sheets/api/quickstart/python">Google Sheets API documentation</a> was used as a reference for building this app.

The requirements for running the presented script are:
- <a href="https://www.microsoft.com/pt-br/software-download/">Windows 8.1</a> or higher, preferentially the <a href="https://www.microsoft.com/pt-br/software-download/">Windows 10</a> version;
- <a href="https://www.python.org/downloads/">Python 3</a> installed, preferentially the most recent version;
- Download the uploaded files in the repository: <a href="./credentials.json">**credentials.json**</a>; <a href="./solution_script.py">**solution_script.py**</a>, and <a href="./requirements.txt">**requirements.txt**</a>;
- Download the libraries presented on the file <a href="./requirements.txt">**requirements.txt**</a>.

# Setting up the environment

To modify Google Sheets spreadsheets using Python, the Google Sheets API was required. Through the four steps below:

- **Step 1**: Create a project on <a href="https://console.cloud.google.com/">*Google Cloud*</a>;
- **Step 2**: Select the created project, and activate the necessary APIs;
- **Step 3**: Configure the *OAuth* permission screen;
- **Step 4**: Authorize credentials for a computer application.

After this setup, the <a href="./credentials.json">**credentials.json**</a> file was generated, and then the Google client library for Python was installed. For that, the following code was executed at the VSCode terminal:
```
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

# Creating the app
After setting up the environment, Python commands could finally be executed. In the first place, the required libraries were imported:
```python
import os.path
import math
import re
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
```

Next, the API scopes, the spreadsheet, and its cell range to be used were set up:
```python
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of the spreadsheet
SPREADSHEET_ID = "1y1beFCzEiD4xYjzVczTWXW4TDkemwNolMb-aPptoJKs"
RANGE_NAME = "engenharia_de_software!C4:F27"

# The cell range where the total number of classes is located
CLASSES_SPREADSHEET_RANGE = "A2:H2"

# The missing spreadsheet data to update
DATA_TO_UPDATE = "engenharia_de_software!G4:H27"
```

Then, the following code was executed to update the blank spreadsheet values:
```python
def log(message):
  '''
  Returns a log line.
  '''
  now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  return print('[{now}] {message}'.format(now=now, message=message))

def get_auth_credentials_token():
  '''
  Authenticates the user's Google account, and generates a json file named "token".
  '''
  # Start with empty creds
  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  if not creds:
    flow = InstalledAppFlow.from_client_secrets_file(
      "credentials.json", SCOPES
    )
    creds = flow.run_local_server(port=0)

  # If there are no (valid) credentials available, let the user log in.
  if creds and not creds.valid:
    if creds.expired and creds.refresh_token:
      creds.refresh(Request())
  
  # Save the credentials for the next run
  with open("token.json", "w") as token:
    token.write(creds.to_json())

  return creds

def school_status(school_absences, average_grade, total_classes):
  '''
  Returns the status of a student based on the number of absences and average grade.
  '''
  try:
    if school_absences > (total_classes * 0.25):
      situation = "Reprovado por Falta"
    else:
      if average_grade < 50:
        situation = "Reprovado por Nota"
      elif average_grade >= 50 and average_grade < 70:
        situation = "Exame Final"
      elif average_grade >= 70:
        situation = "Aprovado"
      
    if situation == "Exame Final":
      # The math.ceil method rounds a number up to the nearest integer
      minimum_grade = math.ceil(100 - average_grade)
      approval_grade = f">= {minimum_grade}"
    else:
      approval_grade = 0
    
    return situation, approval_grade

  except Exception as e:
    log(e)

def get_total_classes_quantity(sheet):
  '''
  Returns the number of classes in the semester.
  '''
  try:
    # Getting the number of classes in the semester
    data_total_classes = sheet.values().get(
      spreadsheetId=SPREADSHEET_ID,
      range=CLASSES_SPREADSHEET_RANGE).execute()
    
    # [0][0] indexing refers to the cell itself - first column of the first row
    total_classes = int( 
      re.search( r'\d+', data_total_classes["values"][0][0] )[0]
    )

    return total_classes

  except Exception as e:
    log(e)
  
def get_students_values(sheet):
  '''
  Returns a matrix that contains the information of each student: absences and exam grades.
  '''
  try:
    students_values = sheet.values().get(
      spreadsheetId=SPREADSHEET_ID,
      range=RANGE_NAME).execute()['values']
      
    return students_values
  
  except Exception as e:
    log(e)

def add_data(students_values, total_classes):
  '''
  Returns a matrix, which contains the data to add to the spreadsheet.
  '''
  try:
    data_to_add = []
    for student_info in students_values:
      absences, *exams = student_info
      absences = int(absences)
      exam_1, exam_2, exam_3 = [float(exam) for exam in exams]
      average_grade = (exam_1 + exam_2 + exam_3) / 3
      situation, approval_grade = school_status(absences, average_grade, total_classes)
      data_to_add.append([situation, approval_grade])

    return data_to_add
  
  except Exception as e:
    log(e)

def update_data(sheet, data_to_add):
  '''
  Updates the spreadsheet data - cell range G4:H27.
  '''
  try:
    updated_data = sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=DATA_TO_UPDATE, valueInputOption="USER_ENTERED",
        body={"values": data_to_add}).execute()
    
    return updated_data

  except Exception as e:
    log(e)
  
def main():
  log('Authenticating with Google...')
  creds = get_auth_credentials_token()
  log('Authentication successful!')

  # The following lines manipulate the data in Google Sheets
  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    log('Calling the Google Sheets API...')
    sheet = service.spreadsheets()
    log('API sucessfully called!')

    # Getting the number of classes in the semester
    log('Getting the required informations in the spreadsheet...')
    total_classes = get_total_classes_quantity(sheet)

    # Getting the students information
    students_values = get_students_values(sheet)
    log('Informations acquired! Now it is time to update some data...')

    # Defining the data to be added
    data_to_add = add_data(students_values, total_classes)

    # Updating data
    update_data(sheet, data_to_add)
    log('Data updated!')

  except HttpError as e:
    log(e)

# Executing as a script
if __name__ == "__main__":
  main()
```

# Running the app

The following steps must be followed for running the app:
- **Step 1**: download Python, if you haven't it installed on your computer;
- **Step 2**: open the following <a href="https://docs.google.com/spreadsheets/d/1y1beFCzEiD4xYjzVczTWXW4TDkemwNolMb-aPptoJKs/edit#gid=0">spreadsheet</a>;
- **Step 3**: create a folder on your personal computer, download the files <a href="./solution_script.py">**solution_script.py**</a>; <a href="./credentials.json">**credentials.json**</a>, and <a href="./requirements.txt">**requirements.txt**</a>, move them to the created folder;
- **Step 4**: open the folder previously created, right-click the file <a href="./solution_script.py">**solution_script.py**</a>, click on "properties" and copy the file location;
- **Step 5**: open the command line (cmd) using the key shortcut `win + R`. Write `cd`, add a space, and enter the text copied from the previous step;
- **Step 6**: now, paste the code "`pip install -r requirements.txt`", and press enter. After the installation, enter "`python solution_script.py`";
- **Step 7**: in the opened window, select your Google account, and a security window will be shown. In this window, click "Advanced", and click "Access case_tunts.rocks (not secure)" to open the authentication window. Now click "continue" to authenticate the app;
- **Step 8**: after the authentication, a file named `token.json` will be generated in the folder you previously created, and a window with the message "The authentication flow has completed. You may close this window." will appear. Close this window, and notice the changes in the spreadsheet cell range **(G4:H27)** values;
- **Step 9** (optional): type the code `pip uninstall -r requirements.txt` to uninstall the libraries previously installed for running this application.
