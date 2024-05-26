import openpyxl
import csv
from collections import Counter
from datetime import datetime

def checkValidPassword(password):

    # Requirement 1: Password length should be at least eight characters
    if len(password) < 8:
        return False
    
    # Requirement 2: Password should contain at least one number or symbol
    if not any(char.isdigit() or char in "!@#$%^&*()-_=+[]{}|;:'\"<>,.?/~`" for char in password):
        return False
    
    # Requirement 3: Password should contain both uppercase and lowercase letters
    if not any(char.isupper() for char in password) or not any(char.islower() for char in password):
        return False
    
    return True

def getUnproblematicPasswords(filePath, compromisedPasswordFile):

    try:
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(filePath)
        
        # Get the active worksheet
        sheet = workbook.active
        
        # Create a set to store passwords from the dictionary file
        dictionaryPasswords = set()
        
        # Load compromised passwords from the CSV file
        with open(compromisedPasswordFile, 'r', encoding='latin-1') as f:
            reader = csv.reader(f)
            for row in reader:
                dictionaryPasswords.add(row[1])  # Second Column is PasswordHeader
        
        # Create a dictionary to store valid passwords for each user
        validPasswords = {}
        
        # Create a list to store all passwords
        allPasswords = []
        
        # Get the current date
        currentDate = datetime(2022, 5, 4)
        
        # Iterate through rows starting from the second row
        for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
            username = row[0]
            password = row[4]
            passwordChangeDate = row[5]
            
            # Check if password meets all criteria and password change date is within the last 90 days
            if checkValidPassword(password) and passwordChangeDate and password not in dictionaryPasswords:
                # Calculate the difference in days between current date and password change date
                daysSinceLastChange = (currentDate - passwordChangeDate).days
                if daysSinceLastChange <= 90:
                    allPasswords.append(password)
                    
                    # Check if the user already has a password and if the new password is unique
                    if username in validPasswords:
                        validPasswords[username].append((password, index))  # Storing password along with its index
                    else:
                        validPasswords[username] = [(password, index)]

        # Count occurrences of each password
        passwordCounts = Counter(allPasswords)
        
        # Output header as per instructions
        print(f"{'Record':<16}{'UserName':<16}{'Password'}")
        
        # List to store sorted rows
        sorted_rows = []

        # Display valid passwords for each user
        for username, passwords in sorted(validPasswords.items(), key=lambda x: x[0]):  # Sorting by username
            for password, index in passwords:
                # Check if the password appears only once in the source data
                if passwordCounts[password] == 1:
                    sorted_rows.append((index, username, password))

        # Print sorted rows
        for row in sorted_rows:
            print(f"{row[0]:<16}{row[1]:<16}{row[2]}")
        
        # Display the count of valid passwords
        validPasswordCounts = sum(1 for password in allPasswords if passwordCounts[password] == 1)
        print(f"Total valid passwords: {validPasswordCounts}")
        
    except FileNotFoundError:
        print(f"File {filePath} or {compromisedPasswordFile} not found.")

# Problem 5: Get unproblematic Passwords
getUnproblematicPasswords("Cybersecurity_case_studies_LibertyDataSystems_Sample.xlsx", 
                  "Cybersecurity_case_studies_LibertyDataSystems_PasswordDictionary.csv")
