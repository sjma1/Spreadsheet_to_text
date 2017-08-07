import openpyxl, os

def get_spreadsheet_location():
    while True:
        spreadsheet = input("Enter the path of the .xlsx file you want to use: ")
        if os.path.isfile(spreadsheet):
            return spreadsheet
        print('INVALID PATH\n')
        
def get_new_path():
    while True:
        directory = input("Enter the path of the directory the .txt files will be placed: ")
        if os.path.isdir(directory):
            return directory
        print('INVALID PATH\n')
        
def get_read_choice():
    while True:
        choice = input("Read the .xlsx file by rows or by columns (r/c): ")
        if choice[0].lower() == 'r' or choice[0].lower() == 'c':
            return choice[0]
        print('INVALID CHOICE, ENTER EITHER \'r\' or \'c\'\n')
        
def write_xlsx_to_txt(spreadsheet_location, txt_directory, read_choice):
    workbook = openpyxl.load_workbook(spreadsheet_location)

    for sheet in workbook.worksheets:
        os.chdir(os.makedirs(txt_directory + '\\' + sheet.title))
        
        if read_choice == 'r':
            sections = sheet.rows
            filename = 'row'
        else:
            sections =  sheet.columns
            filename = 'column'
        for index in range(len(sections)):
            new_file = open(filename + str(index + 1) + '.txt', 'w')
            
        

