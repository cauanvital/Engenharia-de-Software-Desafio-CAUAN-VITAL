# Imports the required extensions to execute the script
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

    # Checks if token and credentials json files are configured correctly to grant
    # write access to the spreadsheet
    # This section can create a new access token if necessary, but this
    # requires granting direct access to an email account on Google Cloud
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
        # Builds the variable to call the spreadsheet related with the credential
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()

        # Creates the variables to get the spreadsheet values and add the new values in the future
        exams = sheets.values().get(spreadsheetId=SHEET_ID,
                                    range=f"challenge_sheet!C4:F27").execute().get("values", [])
        new_list_values = []
        process_number = 4
        
        # Repeat this section to each row of the spreadsheet
        for row in exams:
            # Creates the average variable and print on screen the current column in processing
            # Also defines a list to store the new values to be added to the spreadsheet
            row_new_values = ["", ""]
            m = (int(row[1]) + int(row[2]) + int(row[3])) / 3
            print(f"Processing row {process_number} of the spreadsheet")
            process_number += 1

            # Checks the student's current average to assign the student the
            # correct situation status
            if int(row[0]) > 15:
                new_situation = "Reprovado por Falta"
            elif m < 50:
                new_situation = "Reprovado por Nota"
            elif 70 > m >= 50:
                new_situation = "Exame Final"
            else:
                new_situation = "Aprovado"

            # Adds the new_situation value to the list to update the final list
            row_new_values[0] = new_situation

            # Checks the student assigned situation status
            # and applies the correct necessary score to be approved
            if new_situation == "Exame Final":
                naf = round(100 - m + 0.5)
            else:
                naf = 0

            # Update the naf value to the row_new_values variables
            # and assign the data to the final variable
            row_new_values[1] = naf
            new_list_values.append(row_new_values)
        
        # Update the spreadsheet with list generated by the for loop
        sheets.values().update(spreadsheetId=SHEET_ID,
                               range=f"challenge_sheet!G4:H27",
                               valueInputOption="USER_ENTERED",
                               body={"majorDimension": "ROWS", "values": new_list_values}).execute()
        print("Update complete!")
    # Return an error if something goes wrong
    except HttpError as error:
        print(error)


# Runs the function to complete the challenge ;)
if __name__ == "__main__":
    main()
