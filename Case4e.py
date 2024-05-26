import openpyxl

def countNeverChangedPassword(filePath):
    
    try:
        # Load the workbook passed in as arg
        workbook = openpyxl.load_workbook(filePath)
       
        # Get the active worksheet
        sheet = workbook.active
        
        # Create a list to store usernames and passwords that match the criteria
        passwordNeverChanged = []

        # Iterate through rows starting from the second row
        for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
            username = row[0]
            password = row[4]
            passwordChangeDate = row[5]

            # Check if PwdChangeDate is null or empty
            if not passwordChangeDate:
                passwordNeverChanged.append((username, password, index))

        # Sort passwordNeverChanged by username in ascending order
        passwordNeverChanged.sort(key=lambda x: x[0])
        
        # Output header as per instructions
        print(f"{'Record':<16}{'UserName':<16}Password")

        # Display usernames and passwords that match the criteria
        for username, password, index in passwordNeverChanged:
            print(f"{index:<16}{username:<16}{password}")

        # Display the count of users who have never changed their passwords
        print(f"Total users who have never changed their password: {len(passwordNeverChanged)}")
        
    except FileNotFoundError:
        print(f"File {filePath} not found. We assume an XLSX file which I wish was a CSV")

# Problem 4e get Passwords that have not been Changed
countNeverChangedPassword("Cybersecurity_case_studies_LibertyDataSystems_Sample.xlsx")
