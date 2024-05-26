import openpyxl

def countPasswordsWithLettersOnly(filePath):
    
    try:
        # Load the workbook passed in as arg
        workbook = openpyxl.load_workbook(filePath)
       
        # Get the active worksheet
        sheet = workbook.active
        
        # Counter for passwords with letters only
        lettersOnlyCount = 0
        
        # Output header as per instructions
        print(f"{'Record':<16}{'UserName':<16}{'Password'}")
        
         # List to store rows
        rows = []

        # Iterate through rows but avoid header (so record number excludes header)
        for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
            username = row[0]
            password = row[4]

            # Check if password contains only letters
            if password.isalpha():
                lettersOnlyCount += 1
                rows.append((index, username, password))

        # Sort rows by username in ascending order
        sorted_rows = sorted(rows, key=lambda x: x[1])

        # Print sorted rows
        for row in sorted_rows:
            print(f"{row[0]: <16}{row[1]: <16}{row[2]}")

        # Print total count of passwords with letters only
        print('\nTotal passwords containing only letters:', lettersOnlyCount, "\n")
        
    except FileNotFoundError:
        print(f"File {filePath} not found. We assume an XLSX file which I wish was a CSV")

# Problem 4b get Passwords with only Letters:
countPasswordsWithLettersOnly("Cybersecurity_case_studies_LibertyDataSystems_Sample.xlsx")