import openpyxl as xl
import os

# gets the sheet and saves the entries in row 2 in a text file
def getHeaders(fileName, sheetName):

    # workbook is the file, sheet is just the page in the file
    workbook = xl.load_workbook(fileName)
    sheet = workbook[sheetName]

    headers = "lookupDict = {\n"

    # iterate through cells in row 2 (that's where the headers are)
    for cell in sheet[2]:

        if cell.value == "Requirement - Would you like to enter another response?":
            break

        header = ""

        # put in assumed headers
        if "GEG" in cell.value:
            header = "Goal"
        elif "Requirement - Select the core course, foundation, or skill & perspective for your data." in cell.value:
            header = "Requirement"
        elif "Select the course you taught" in cell.value:
            header = "Course"
        elif "type of assessment" in cell.value:
            header = "Assessment Type"
        elif "acceptable achievement" in cell.value:
            header = "Met"
        if "unacceptable achievement" in cell.value:
            header = "Not Met"
        elif "First Name" in cell.value:
            header = "First Name"
        elif "Last Name" in cell.value:
            header = "Last Name"
        elif "Email" in cell.value:
            header = "Email"
        
        if header == "":
            # print(cell.value)
            continue

        headers += "\t\"" + cell.value + "\": \"" + header + "\",\n"


    # remove the last comma and close the dictionary
    headers = headers[:-2]
    headers += "\n}"

    # delete the file if it already exists
    if os.path.exists("headers.txt"):
        os.remove("headers.txt")

    # save to headers.txt
    with open("headers.txt", "w") as text_file:
        text_file.write(headers)

    return

if __name__ == "__main__":

    getHeaders("test.xlsx", "Raw")

