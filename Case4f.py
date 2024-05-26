import openpyxl
from datetime import datetime

def countPasswordsNotChangedLast90Days(filePath):
    
    try:
        # Load the workbook passed in as arg
        workbook = openpyxl.load_workbook(filePath)
       
        # Get the active worksheet
        sheet = workbook.active
        
        # Create a list to store usernames and passwords that match the criteria
        passwordsNotChangedLast90Days = []

        # Current date
        currentDate = datetime(2022, 5, 4)

        # Iterate through rows starting from the second row
        for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
            username = row[0]
            password = row[4]
            passwordChangeDate = row[5]

            # Check if passwordChangeDate is not null or empty and is more than 90 days ago
            if passwordChangeDate:
                daysSinceLastChange = (currentDate - passwordChangeDate).days
                if daysSinceLastChange > 90:
                    passwordsNotChangedLast90Days.append((username, password, daysSinceLastChange, index))

        # Sort passwordsNotChangedLast90Days by username in ascending order
        # We only had one result but I still wrote it like this in case we have more
        passwordsNotChangedLast90Days.sort(key=lambda x: x[0])

        # Output header as per instructions
        print(f"{'Record':<16}{'UserName':<16}{'Password':<16}{'TimeSincePwdChange'}")

        # Display usernames, passwords, and days since last change that match the criteria
        for username, password, daysSinceLastChange, index in passwordsNotChangedLast90Days:
            print(f"{index:<16}{username:<16}{password:<16}{daysSinceLastChange}")

        # Display the count of users who have not changed their passwords in the last 90 days
        print(f"Total users who have not changed their password in the last 90 days: {len(passwordsNotChangedLast90Days)}")
        
    except FileNotFoundError:
        print(f"File {filePath} not found. We assume an XLSX file which I wish was a CSV")

# Problem 4f: Get passwords that have not been changed in the last 90 days
countPasswordsNotChangedLast90Days("Cybersecurity_case_studies_LibertyDataSystems_Sample.xlsx")
