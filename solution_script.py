# Required libraries
import os.path
import math
import re
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of the spreadsheet
SPREADSHEET_ID = "1y1beFCzEiD4xYjzVczTWXW4TDkemwNolMb-aPptoJKs"
RANGE_NAME = "engenharia_de_software!C4:F27"

# The cell range where the total number of classes is located
CLASSES_SPREADSHEET_RANGE = "A2:H2"

# The missing spreadsheet data to update
DATA_TO_UPDATE = "engenharia_de_software!G4:H27"

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
  Returns a matrix that contains the informations of each student: absences and exam grades.
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
