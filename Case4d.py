import openpyxl

def countDuplicatePasswords(filePath):
    
    try:
        # Load the workbook passed in as arg
        workbook = openpyxl.load_workbook(filePath)
       
        # Get the active worksheet
        sheet = workbook.active
        
        # Create a dictionary to store usernames and passwords
        passwordDict = {}

        # Iterate through rows starting from the second row
        for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
            username = row[0]
            password = row[4]

            if password in passwordDict:
                passwordDict[password].append((username, index))
            else:
                passwordDict[password] = [(username, index)]

        # Output header as per instructions
        print(f"{'Record':<16}{'UserName':<16}Password")

        # List to store sorted rows
        sorted_rows = []

        # Iterate through the password dictionary
        for password, userRow in sorted(passwordDict.items()):
            
            # Check if there are more than one username with the same password
            if len(userRow) > 1:
                # Sort the usernames by initial row index
                sortedUserRows = sorted(userRow, key=lambda x: x[0])  # Sorting by username

                # Add sorted rows to the list
                for username, index in sortedUserRows:
                    sorted_rows.append((index, username, password))

        # Sort sorted rows by username in ascending order
        sorted_rows = sorted(sorted_rows, key=lambda x: x[1])

        # Print sorted rows
        for row in sorted_rows:
            print(f"{row[0]:<16}{row[1]:<16}{row[2]}")

        # Display the count of users who have duplicate passwords
        print(f"Total users who have duplicate passwords: {len(sorted_rows)}")
        
    except FileNotFoundError:
        print(f"File {filePath} not found. We assume an XLSX file which I wish was a CSV")

# Problem 4d get Duplicate Passwords
countDuplicatePasswords("Cybersecurity_case_studies_LibertyDataSystems_Sample.xlsx")
