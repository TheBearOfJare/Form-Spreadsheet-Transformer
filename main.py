import openpyxl as xl

# handy colors to differentiate print statements in the terminal
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def getSheet(fileName, sheetName):

    # workbook is the file, sheet is just the page in the file
    workbook = xl.load_workbook(fileName)
    sheet = workbook[sheetName]

    return sheet

# this is a convenience function that returns the what column name a response for a question should go in (ex. "Requirement - Select the core course, foundation, or skill & perspective for your data." returns "Requirement") The purpose is to make the code more readable and to make it easier to maintain, as changing the form and/or questions should genrally necessitate changing this function only and nothing else
def getDataHeader(QuestionDescriptor):

    lookupDict = {
        "Start Date": "A",
        "End Date": "B",
        "Response Type": "C",
        "IP Address": "D",
        "Progress": "E",
        "Duration (in seconds)": "F",
        "Finished": "G",
        "Recorded Date": "H",
        "Response ID": "I",
        "Recipient Last Name": "Last Name",
        "Recipient First Name": "First Name",
        "Recipient Email": "Email",
        "External Data Reference": "M",
        "Location Latitude": "N",
        "Location Longitude": "O",
        "Distribution Channel": "P",
        "User Language": "Q",
        "Please complete this form to submit General Education Program data. Let's begin with your name. - First Name": "First Name",
        "Please complete this form to submit General Education Program data. Let's begin with your name. - Last Name": "Last Name",
        "Please complete this form to submit General Education Program data. Let's begin with your name. - BC Email Address": "Email",
        "Requirement - Select the core course, foundation, or skill & perspective for your data.": "Requirement",
        "Requirement - THEO-1100 is preselected as the course you taught that fulfills the Introduction to Theology requirement.": "V",
        "Requirement - The General Education Goal (GEG) for Introduction to Theology is shown below.": "Goal",
        "Requirement - Select the course you taught that fulfills the Faith requirement.": "Course",
        "Requirement - Select the General Education Goal(s) (GEG) for Faith (F) you are taught and assessed in [QID65-ChoiceGroup-SelectedChoices].": "Goal",
        "Requirement - Select the course you taught that fulfills the Mathematical Reasoning (MR) requirement.": "Course",
        "Requirement - Select the General Education Goal(s) (GEG) for Mathematical Reasoning (MR) you taught and assessed in [QID66-ChoiceGroup-SelectedChoices].": "Goal",
        "Requirement - Select the course you taught that fulfills the Scientific Method (SM)  requirement.": "Course",
        "Requirement - Select the General Education Goal(s) (GEG) for Scientific Method (SM) you taught and assessed in [QID67-ChoiceGroup-SelectedChoices].": "Goal",
        "Requirement - Select the course you taught that fulfills the Understanding the Natural World (NW) requirement.": "Course",
        "Requirement - Select the General Education Goal(s) (GEG) for Understanding the Natural World (NW) you taught and assessed in [QID68-ChoiceGroup-SelectedChoices].": "Goal",
        "What type of assessment did you use to gather data? - Requirement - What type of assessment did you use to gather data? - Selected Choice": "Assessment Type",
        "What type of assessment did you use to gather data? - Requirement - Multiple/Other (please describe) - Text": "Assessment Type",
        "What were your assessment results? - Requirement - How many students demonstrated acceptable achievement?": "Met",
        "What were your assessment results? - Requirement - How many students demonstrated unacceptable achievement?": "Not Met"
    }

    try: 
        return lookupDict[QuestionDescriptor]
    except KeyError:
        print(bcolors.FAIL + "No header found for question: " + QuestionDescriptor + bcolors.ENDC)
        return None

def transform(sheet):

    # data is a list of dictionarys, where each outer list is a row in the spreadsheet, and the dictionary is the relationship between the headers and the data
    data = []

    # rowCount represents the the number of already existing rows in the final, transformed spreadsheet, which is not nessicarily the same as the current row in the raw spreadsheet

    rowCount = 0

    # iterate through each row (past the first two because those are headers). Each row contains all the data for one person.
    for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column, values_only=True):

        # check that there is data in the first cell of the row
        if row[0] == None:
            continue

        # iterate through each cell in the row
        i = 0
        for cell in row:

            # check that the cell is not empty
            if cell == None:
                continue

            # get the header for the cell with getDataHeader() and the value of the cell in row 2 in cell's column
            label = sheet.cell(row=2, column=i).value
            header = getDataHeader(label)

            # check that we know where the data goes
            if header == None:
                continue
            
            # check that there isn't already data in data[row.index][header]. If there is, increment rowCount to start putting the data in the next row. 
            if data[rowCount][header] != None:
                rowCount += 1

            data[rowCount][header] = cell

            i += 1

    return data

def main():
    
    # the raw data has each submission from the same user in one row, and we want to split it into individual rows for each response, even from the same person. Users who submit multiple responses will generate multiple rows in the final spreadsheet.

    # get the sheet
    sheet = getSheet("test.xlsx", "Raw")

    # headers is a list of all the headers in the spreadsheet, data is a list of dictionarys, where each outer list is a row in the spreadsheet, and the dictionary is the relationship between the headers and the data
    headers = ["First Name", "Last Name", "Email", "Requirement", "Course", "Goal", "Assessment Type", "Met", "Not Met"]
    data = transform(sheet=sheet)
    
    # write the data to a new spreadsheet

    # create the new spreadsheet
    transformed = xl.Workbook()
    sheet = transformed.create_sheet(title="Transformed")

    # make the headers
    for i in range(len(headers)):
        sheet.cell(row=1, column=i+1).value = headers[i]

    # add the data
    for i in range(len(data)):
        for j in range(len(headers)):
            sheet.cell(row=i+2, column=j+1).value = data[i][headers[j]]

    # save the spreadsheet and we're done
    transformed.save(f"transformed.xlsx")

    print(bcolors.OKGREEN + "Done!" + bcolors.ENDC)


if __name__ == "__main__":

    main()