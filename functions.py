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
        new_path = txt_directory + '\\' + sheet.title
        os.makedirs(new_path)
        os.chdir(new_path)
        
        if read_choice == 'r':
            sections = []
            for r in sheet.rows:
                sections.append(r)
            filename = 'row'
        else:
            sections =  []
            for c in sheet.columns:
                sections.append(c)
            filename = 'column'
            
        for index in range(len(sections)):
            new_file = open(filename + str(index + 1) + '.txt', 'w')
            for item in sections[index]:
                new_file.write(str(item.value) + '\t')
            new_file.close()
        

