import openpyxl

def countShortPasswords(filePath):
    
    try:
        # Load the workbook passed in as arg
        workbook = openpyxl.load_workbook(filePath)
       
        # Get the active worksheet
        sheet = workbook.active
        
        # Counter for short passwords
        shortPasswordNum = 0
        
        # Output our header per instructions
        print(f"{'Record':<16}{'UserName':<16}{'Password':<16}{'Length'}")

        # List to store rows
        rows = []

        # Iterate through rows but avoid header (so record number excludes header)
        for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
            
            passwordLength = len(row[4]) # The password is in the 5th column but we use 0 indexes
            username = row[0]
            password = row[4]

            if passwordLength < 8:
                shortPasswordNum += 1 # Add to our short passwords count
                rows.append((index, username, password, passwordLength))

        # Sort rows by username in ascending order
        sorted_rows = sorted(rows, key=lambda x: x[1])

        # Print sorted rows
        for row in sorted_rows:
            print(f"{row[0]: <16}{row[1]: <16}{row[2]: <16}{row[3]}")

        # Print total count of short passwords to validate
        print('\nTotal passwords less than eight characters:', shortPasswordNum,"\n")
        
    except FileNotFoundError:
        print(f"File {filePath} not found. We assume a XLSX file which I wish was a CSV")

# Problem 4a get Short Passwords
countShortPasswords("Cybersecurity_case_studies_LibertyDataSystems_Sample.xlsx")  # Make sure the file is in current dir or path
