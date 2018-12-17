#Final project rewrite
import openpyxl
import snaps
import cherrypy
wb = openpyxl.load_workbook('Program Source Spreadsheet.xlsx')
#Imports openpyxl and loads the Excel sheet
def cell(sColumn, iRow, iSheet=1):
     return wb ['Sheet' + str (iSheet)][sColumn + str(iRow)].value
#Defining a function to be used later. Placed up here to clear up the list of functions later. 
print ('''Welcome to the Shakespeare Library's Book Recommendation Program!''')
print ('''Are you ready to try some diverse genre picks?''')
#Creates a little introduction to the program for new users. Library name clearly made up

def inputQuery (sMessageToUser):
     #Defining a function to input the input string at the bottom of the document
    while True:
        sQuery = input (sMessageToUser)
        if (sQuery == ''):
            print ('\nYou did not enter anything.')
            print ('Please enter something or close the tab.')
            continue
        if sQuery.isnumeric():
            print ('\nYou entered ' + sQuery + '. This is a number.')
            print ('Please enter a genre')
            continue
        if not sQuery in ['Fantasy', 'Science Fiction', 'Romance', 'Mystery', 'Horror']:
            print ('\nYou entered ' + sQuery + ', which the program does not recognize as an option.')
            print ('You may want to check if the capitalization matches the given choices.')
            print ('Please try again.')
            continue
        return sQuery
#Conditionals with the most common wrong answers
def inputPrintSpec():
    return ['B','D','C','F','E']
#Function to return specific columns in a specific order
#columns are out of order because the excel sheet isn't designed for user examination, but for library use
def printRecord(iRow):
    print()
    for sCol in inputPrintSpec():
        print (cell(sCol, 1).ljust(17)+':  ' + str(cell(sCol, iRow)))
#Function that prints the columns collected above if the function below is successful
def search (sQueryString):
    iFound = 0
    print ('\nSearching for diverse ' + sQueryString + ' books \n')
    for iRow in range(2, 100000):
        if cell('A', iRow) == None:
            print ('\n' + str(iFound) + ' records found')
            break
        if cell('A', iRow).startswith(sQueryString):
            printRecord(iRow)
            iFound = iFound+1
        continue
#Function that searches for the string in the Excel document and returns what it gets
sAnotherSearch = 'Yes'

while True:

    if sAnotherSearch in ['N', 'n', 'NO', 'No', 'no', 'PLEASE STOP', 'NOPE', 'TERMINATE', 'please stop']:
        print ('\nGood-bye')
        break
    if sAnotherSearch in ['Y', 'y', 'YES', 'Yes', 'yes']:
        sCorrectQuery = inputQuery ('What genre do you want to try? Options are Fantasy, Science Fiction, Romance, Mystery, and Horror:  ')
        search(sCorrectQuery)
    else:
        print ('Your answer is ' + sAnotherSearch + '. Please answer yes or no.')
    sAnotherSearch = input ('\nWould you like to search again? ')
#While loop with the input that actually starts off the user interaction. Is way down here because it utilizes the search function defined above it
