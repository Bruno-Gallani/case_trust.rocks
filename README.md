The solution was all written in *Python*, and the official <a href="https://developers.google.com/sheets/api/quickstart/python">Google Sheets API documentation</a> was used as reference for building this app.

The requirements for running the presented script are:
- <a href="https://www.python.org/downloads/">Python</a> installed, preferentially the most recent version;
- Download the files uploaded in the repository: <a href="https://github.com/Bruno-Gallani/case_trust.rocks/blob/main/credentials.json">**credentials.py**</a> and <a href="https://github.com/Bruno-Gallani/case_trust.rocks/blob/main/script.py">**script.py**</a>.

# Setting up the environment

In order to modify Google Sheets spreadsheets using Python, I needed to use the Google Sheets API. But, for this, I followed the four steps below:

- **Step 1**: Create a project on <a href="https://console.cloud.google.com/">*Google Cloud*</a>;
- **Step 2**: Select the created project, and activate the necessary APIs;
- **Step 3**: Configure the *OAuth* permission screen;
- **Step 4**: Authorize credentials for a computer application.

After this setup, the <a href="https://github.com/Bruno-Gallani/case_trust.rocks/blob/main/credentials.json">**credentials.json**</a> file was generated, and then I installed the Google client library for Python. For that, I executed the following code at the VSCode terminal:
```
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

# Creating the app
After setting up the environment, we can, finally, execute commands in Python. In first place, we import the required libraries:
```python
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
```

Next, we set up the API scopes, the spreadsheet and its cells range to be used:
```python
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of the spreadsheet
SPREADSHEET_ID = "1y1beFCzEiD4xYjzVczTWXW4TDkemwNolMb-aPptoJKs"
RANGE_NAME = "engenharia_de_software!A4:H27"

# The missing spreadsheet data to update
DATA_TO_UPDATE = "engenharia_de_software!G4:H27"
```

Now, we execute the following code to update the blank spreadsheet values:
```python
def school_status(school_absences, average_grade, total_classes):

  if school_absences > (total_classes * 0.25):
    situation = "Reprovado por Falta"
  else:
    if average_grade < 5:
      situation = "Reprovado por Nota"
    elif average_grade >= 5 and average_grade < 7:
      situation = "Exame Final"
    elif average_grade >= 7:
      situation = "Aprovado"
    
  if situation == "Exame Final":
    minimum_grade = 10 - average_grade
    approval_grade = f"naf >= {minimum_grade:.2f}"
  else:
    approval_grade = 0
  
  return situation, approval_grade

def main():

  # The following lines let the user access the application

  # Start with empty creds
  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  # The following lines manipulate the data in Google Sheets
  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    # Getting the number of classes in the semester
    data_total_classes = sheet.values().get(
      spreadsheetId=SPREADSHEET_ID,
      range="A2:H2").execute()
    total_classes = int(
      data_total_classes["values"][0][0].replace(
      "Total de aulas no semestre: ", ""
    ))

    # Getting the students information
    students_data = sheet.values().get(
      spreadsheetId=SPREADSHEET_ID,
      range=RANGE_NAME).execute()
    students_values = students_data['values']

    # Adding data
    data_to_add = []

    for linha in students_values:
      absences = int(linha[2])
      exam_1, exam_2, exam_3 = float(linha[3]) / 10, float(linha[4]) / 10, float(linha[5]) / 10
      average_grade = (exam_1 + exam_2 + exam_3) / 3
      situation, approval_grade = school_status(absences, average_grade, total_classes)
      data_to_add.append([situation, approval_grade])

    # Updating the data
    sheet.values().update(
      spreadsheetId=SPREADSHEET_ID,
      range=DATA_TO_UPDATE, valueInputOption="USER_ENTERED",
      body={"values": data_to_add}).execute()

  except HttpError as err:
    print(err)

# Executing as a script
if __name__ == "__main__":
  main()
```

# Running the app

To run the app, the following steps must be followed:
- **Step 1**: download Python, if you haven't it installed in your computer;
- **Step 2**: open the link to the <a href="https://docs.google.com/spreadsheets/d/1y1beFCzEiD4xYjzVczTWXW4TDkemwNolMb-aPptoJKs/edit#gid=0">spreadsheet</a>;
- **Step 3**: create a folder on your personal computer, with the archives **script.py** and **credentials.py**;
- **Step 4**: right click on the folder previously created, click on "properties" and copy the local to the file;
- **Step 5**: open the command line (cmd), using the key shorcut `win + R`. Write `cd ` and paste the text copied on the previous step;
- **Step 6**: now, write this code "`pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib`". And then, after the installation, write "`python script.py`";
- **Step 7**: in the opened window, select your google account to do the authentication for the app;
- **Step 8**: after the authentication, close the window, and run again the piece of code "`python script.py`". Notice the changes on the spreadsheet values. 
