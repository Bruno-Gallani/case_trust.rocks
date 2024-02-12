# Required libraries
import os.path
import math
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of the spreadsheet
SPREADSHEET_ID = "1y1beFCzEiD4xYjzVczTWXW4TDkemwNolMb-aPptoJKs"
RANGE_NAME = "engenharia_de_software!A4:F27"

# The missing spreadsheet data to update
DATA_TO_UPDATE = "engenharia_de_software!G4:H27"

def school_status(school_absences, average_grade, total_classes):

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
    approval_grade = f"naf >= {minimum_grade}"
  else:
    approval_grade = 0
  
  return situation, approval_grade

def main():

  # The following lines let the user access the application

  # Start with empty creds
  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
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

    for student_info in students_values:
      absences = int(student_info[2])
      exam_1, exam_2, exam_3 = float(student_info[3]), float(student_info[4]), float(student_info[5])
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
  print("Data updated!")
