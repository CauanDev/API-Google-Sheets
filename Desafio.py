import os.path
import re
import math
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
end = 27
# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1pWrgNznJTFDDZtGnCytzYFuMuHByvZaA0YjJVTxYD-4"
SAMPLE_RANGE_NAME = f"engenharia_de_software!B4:H{end}"

def main():
    # Open the console.txt file in write mode
    file = open('console.txt','w')
    creds = None
    begin = 4
    
    # Check if the token.json file exists and load credentials
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # If no valid credentials are available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
    
    # Save the updated credentials to token.json
    with open('token.json','w') as token:
        token.write(creds.to_json())
    
    try: 
        # Build the Google Sheets service
        service = build('sheets', 'v4', credentials=creds)
    
        # Get values from the spreadsheet
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        quantity = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                      range="engenharia_de_software!A2:H2").execute() 
        quantity = quantity.get('values', [])
        values = result.get('values', [])
        str = quantity[0][0]
        number = int(re.search(r"\d+", str).group())
        
        # Loop through each row in the spreadsheet
        for row in values:
            name = row[0]
            absences = int(row[1])
            first_grade = int(row[2])
            second_grade = int(row[3])
            third_grade = int(row[4])
            mid = (first_grade + second_grade + third_grade) / 3
            
            # Determine student status based on grades and attendance
            if(mid < 50):                                
                str = "Failed due to low grades"
                str1 = 0
            elif(50 <= mid < 70):
                str = "Final Exam"
                calc = 50 - mid
                calc *= 2
                str1 = math.ceil(mid + calc)
            else:
                str = "Passed"
                str1 = 0
            
            # Check attendance and update student status accordingly
            if((math.ceil(absences * 100) / number) > 25):
                str = "Failed due to low attendance" 
                str1 = 0
            
            # Write information to the console.txt file
            file.write(f"The student {name} is {str}\nFinal Grade: {mid} Rounded Final Grade: {math.ceil(mid)}\n\n")
            
            # Update spreadsheet with the calculated status and rounded final grade
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                           range=f"G{begin}", valueInputOption="USER_ENTERED",
                                           body={'values': [[str]]}).execute()
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                           range=f"H{begin}", valueInputOption="USER_ENTERED",
                                           body={'values': [[str1]]}).execute()
            begin += 1
        
        # Close the console.txt file and print a message
        file.close()
        print("Finished")
    except HttpError as error:
        print(error)

if __name__ == "__main__":
    main()
