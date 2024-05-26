import openpyxl
import csv

def getCompromisedPasswords(filePath, compromisedPasswordsFile):
    
    try:
        # Load the workbook passed in as arg
        workbook = openpyxl.load_workbook(filePath)
       
        # Get the active worksheet
        sheet = workbook.active
        
        # Create a set to store passwords from the compromised passwords file
        compromisedPasswordsSet = set()

        # Load passwords from the compromised passwords file
        with open(compromisedPasswordsFile, 'r', encoding='latin-1') as f:
            reader = csv.DictReader(f)
            for row in reader:
                compromisedPasswordsSet.add(row['PasswordHeader'])

        # Create a list to store compromised passwords along with record, username, and password
        compromisedPasswords = []

        # Iterate through rows starting from the second row
        for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
            username = row[0]
            password = row[4]

            # Check if password is in the compromised passwords set
            if password in compromisedPasswordsSet:
                compromisedPasswords.append((index, username, password))
        
        # Sort sorted rows by username in ascending order
        compromisedPasswords = sorted(compromisedPasswords, key=lambda x: x[1])

        # Output header as per instructions
        print(f"{'Record':<16}{'UserName':<16}{'Password'}")

        # Display compromised passwords along with record, username, and password
        for index, username, password in compromisedPasswords:
            print(f"{index:<16}{username:<16}{password}")

        # Display the count of compromised passwords
        print(f"Total compromised passwords: {len(compromisedPasswords)}")
        
    except FileNotFoundError:
        print(f"File {filePath} or {compromisedPasswordsFile} not found.")

# Problem 5: Get passwords that have been compromised
getCompromisedPasswords("Cybersecurity_case_studies_LibertyDataSystems_Sample.xlsx", 
                        "Cybersecurity_case_studies_LibertyDataSystems_PasswordDictionary.csv")
