# Importing the required extensions to execute the script
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Defines the SCOPE and SHEET_ID variables to specify the adress to reference the spreadsheet
SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]
SHEET_ID = "106kadSqmzl8ufnpe1qFTtiYCUgsWRZcskxGydGCq1Fc"


# Defines the main function to connect the program with the spreadsheet and update the requested data
def main():
    credentials = None

    # Checks if token and credentials json files are configured correctly to grant write access to the spreadsheet
    # This section can create a new access token if necessary, but this requires granting direct access to an email account on Google Cloud
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPE)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPE)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())

    # This entire section attempts to get and update worksheet data unless an error occurs
    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()

        # Repeat this section to each row of the spreadsheet
        for row in range(4, 28):
            # Gets the values from "Faltas", "P1", "P2" e "P3" columns from the current row and assign to exams variable
            # Also creates the average variable and print on screen the current column in processing
            exams = sheets.values().get(spreadsheetId=SHEET_ID, range=f"challenge_sheet!C{row}:F{row}").execute().get("values", [])
            m = (int(exams[0][1]) + int(exams[0][2]) + int(exams[0][3])) / 3
            print(f"Processing row {row}")

            # Checks the student's current average to assign the student the correct situation status
            if int(exams[0][0]) > 15:
                new_situation = "Reprovado por Falta"
            elif m < 50:
                new_situation = "Reprovado por Nota"
            elif 70 > m >= 50:
                new_situation = "Exame Final"
            else:
                new_situation = "Aprovado"

            # Write the new_situation value to the spreadsheet
            sheets.values().update(spreadsheetId=SHEET_ID, range=f"challenge_sheet!G{row}",
                                   valueInputOption="USER_ENTERED", body={"values": [[new_situation]]}).execute()

            # Checks the student assigned situation status and applies the correct necessary score to be approved
            if new_situation == "Exame Final":
                naf = round(100 - m + 0.5)
            else:
                naf = 0

            # Write the naf value to the spreadsheet
            sheets.values().update(spreadsheetId=SHEET_ID, range=f"challenge_sheet!H{row}",
                                   valueInputOption="USER_ENTERED", body={"values": [[naf]]}).execute()
    # Return an error if something goes wrong
    except HttpError as error:
        print(error)


# Runs the function to complete the challenge ;)
if __name__ == "__main__":
    main()

