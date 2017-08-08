#Seong Ma

import functions

'''
This is a program that prompts the user to enter the path of 
a .xlsx file and then converts the spreadsheet into multiple 
text files with each file containing the data from either the
rows or the columns of the .xlsx file
'''

if __name__ == '__main__':
    spreadsheet = functions.get_spreadsheet_location()
    text_directory = functions.get_new_path()
    read_choice = functions.get_read_choice()
    
    functions.write_xlsx_to_txt(spreadsheet, text_directory, read_choice)
    
    print('COMPLETE')