#-----------------------------------------------------------------------------------------------------
# 
# Script Name: LLE-Extraction.py
# Author:      Amy Schaeffer
#
# Version: 3
# Start Date:        9/20/2022
# End Date:          9/22/2022
# Description: Identifies required drawing numbers from a range in a user defined Excel 
#              sheet and parses through a Drawing Set PDF, extracting related
#              pages and compiling them into a new PDF file. Additionally, prints list of extracted
#              page data into separate sheet in excel for further review if needed 

# Version 3: Updated - Removed function paring for the drawing set data due to it increasing runtime in 
#            leiu of manually entering data.
#----------------------------------------------------------------------------------------------------

# Importing xlwings to read and print in Excel
import xlwings as xw
#Importing Path to create full paths for files
from pathlib import Path
#Importing PyPDF2 to split pdf into LLE packages
from PyPDF2 import PdfFileReader, PdfFileWriter
#Importing os
import os
# Importing Shutil to move files to directory
import shutil

# Function to add passed key and value to passed dictionary
def add_values_in_dict(sample_dict, key, value):
    sample_dict[key] = value
    return sample_dict 
#
#  Function to add key if it doesnt already exist, and add list as pair value to key in a dictionary
def add_list_in_dict(sample_dict, key, list_of_values):
    if key not in sample_dict:
        sample_dict[key] = list()
    sample_dict[key].extend(list_of_values)
    return sample_dict

# Function that formats drawing numbers to get rid of any extra characters at the end
def format_drawing_number(drawing_number):
    count = 0
    for i in range(len(drawing_number)):
        char = drawing_number[i]
        if char == "0" or char == "1" or char == "2" or char == "3" or char == "4" or char == "5" or char == "6" or char == "7" or char == "8" or char == "9":
            count = count + 1
        if count == 3:
            drawing_number = drawing_number[0:i+1:]
            break
    return drawing_number

# Function to format page number from drawing log to page number only
def format_e_page_num(page_num):
    count = 0
    for i in range(len(page_num)):
        char = page_num[i]
        if char == " ":
            count = count + 1
        if count == 2:
            page_num = page_num[5:i]
            break
    return page_num

#Function to collect electrical data from the LLE Extractor file
def get_drawing_data(drawing_sheet):
    ret = {}
    # Looping through range between rows 4 and 500 for formatting purpose
    for row in range(4,500):
        # Adding values in columns A and B to dictionary until it hits a blank cell
        cell_value = drawing_sheet[f'A{row}'].value
        if cell_value != None:
            key = drawing_sheet[f'A{row}'].value
            value = drawing_sheet[f'B{row}'].value
            value = format_e_page_num(value)
            value = int(value)
            ret[key] = value
        else:
            break
    return ret

# Function to update Data Details sheet with all key value pairs found in page data dictionary
def print_detailed_data(destination_sheet, dictionary):
    #creating counters for printing data into cells in each column
    key_counter = 4
    value_counter = 4
    # printing each key value into column A in sheet, incrementing counter after each iteration
    for key, value in dictionary.items():
        destination_sheet.range(f'A{key_counter}').options(index=False).value = key
        key_counter = key_counter +1
    #printing each value from list into columns B+, incrementing counter after each iteration
    for values, value in dictionary.items():
        destination_sheet.range(f'B{value_counter}').options(index=False).value = value
        value_counter = value_counter +1  

# Function gathering drawing numbers in B column associated with X values for selected package type
def get_package_req(pac_sheet, package_type, m_row_min, m_row_max, p_row_min,  p_row_max, e_row_min, e_row_max):
    # Dictionary to contain required list of drawing numbers as value and package category from row 20 as key 
    req_drawings = {}
    # Indexing column letters as key value pairs to loop through (excel uses [1] index for column letters)
    col_names = {27:"AA",28:"AB",29:"AC",30:"AD",31:"AE",32:"AF",33:"AG",34:"AH",35:"AI",36:"AJ",37:"AK",38:"AL",39:"AM",40:"AN",41:"AO",42:"AP",43:"AQ",44:"AR",45:"AS"}

    # Nested function taking row range parameter to only gather drawing numbers for matching package type
    def req_drawings_for_package_type(rows):
        # Dictionary to hold package category key and drawing lists as values
        ret = {}
        # Looping through key value pairs in col_names dictionary to loop through columns AA-AS
        for num, col in col_names.items():
            # Assigning list name variable to package category value in row 20
            list_name = pac_sheet[f'{col}20'].value
            # Adding package category as key pair with empty list as value
            ret[list_name] = []
            # Looping through range between passed rows
            for row in range(rows[0], rows[1]+1):
                # Collecting values for each cell in range
                value = pac_sheet[f'{col}{row}'].value
                # Checking if an x is in that cell (capital or not)
                if value == "X" or value == "x":
                    # If there is an x in the cell, assigning the value in the B column of that cell's row to a variable (the drawing number)
                    drawing_num = pac_sheet[f'B{row}'].value
                    drawing_num = format_drawing_number(drawing_num)
                    # Appending that drawing number from the B column into the list of values for that columns package category key
                    ret[list_name].append(drawing_num)
        # creating a new dictionary, only incuding lists with value length greater than 0
        ret = {k: v for k, v in ret.items() if len(v) > 0}
        # Returning dictionary with lists of package categories that have a drawing requirement for the LLE Package
        return ret
    
    # Calling nested function and finding drawing numbers conditional to package type, assigning returned dictionary to req_drawings
    if package_type == "Mechanical":
        req_drawings = req_drawings_for_package_type((m_row_min, m_row_max))
    # Else if package type = Plumbing
    elif package_type == "Plumbing":
        req_drawings = req_drawings_for_package_type((p_row_min, p_row_max))
    # Else if package type = Electrical
    elif package_type == "Electrical":
        req_drawings = req_drawings_for_package_type((e_row_min, e_row_max))
    # Returning dictionary with key value being package category and value being list of drawing numbers, only for package categories with drawing numbers marked "x" for required in their column and range
    return req_drawings 

# Function to print required drawings to excel
def print_req_drawings(sheet, dictionary):
    # Creating counters for printing data into cells in each column
    key_counter = 4
    value_counter = 4
    # Printing each key value into column D in sheet, incrementing counter after each iteration
    for key, value in dictionary.items():
        sheet.range(f'D{key_counter}').options(index=False).value = key
        key_counter = key_counter +1
    #printing each value from list into columns E+, incrementing counter after each iteration
    for values, value in dictionary.items():
        sheet.range(f'E{value_counter}').options(index=False).value = value
        value_counter = value_counter +1  

# Function to locate page numbers of required drawings and put page numbers in a list with LLE package as key, Returns dictionary with LLE Package title as key and list of page numbers as values
def get_req_page_num(sheet, page_data, package_req):
    # dictionary to hold page numbers, list for drawings located and missing drawings
    page_nums = {}
    found_drawings = []
    missing_drawings = []
    # Setting count to 5 for sheet format purposes
    count = 5
    # Looping through key value pairs in required drawings data
    for req_key, req_value in package_req.items():
        # Assigning package title to key and making a list for drawing numbers as value
        page_nums[req_key] = []
        # Looping through drawing data to get page numbers 
        for drawing_key, drawing_value in page_data.items():
            formatted_key = format_drawing_number(drawing_key)
            # If drawing number in drawing data matches drawing number in required drawings, append page number to list value for that package key
            if formatted_key in req_value:
                # Adding to list of drawings located
                found_drawings.append(formatted_key)
                page_nums[req_key].append(drawing_value)
    # Looping through value lists in required drawings
    for req_value in package_req.values():
        for x in req_value:
            value = format_drawing_number(x)
            # If the required drawing number is not located in found drawings list, add it to missing drawings and print to sheet
            if value not in found_drawings:
                if value not in missing_drawings:
                    missing_drawings.append(value)
                    sheet.range(f'M{count}').options(index=False).value = value
                    count = count + 1
    return page_nums

# Fucntion to use page_nums to split file by page numbers in value list, and name the new file the key variable.
def create_files(cwd, pdf, page_nums):
     # Reading file
    pdf = PdfFileReader(pdf)
    # Looping through dictionary, taking required page numbers and writing them to key value pdf
    for k, v in page_nums.items():
    # Looping through dictionary, taking required page numbers and writing them to key value pdf
        #creating pdf writer object
        pdfWriter = PdfFileWriter()
        for page_num in v:
            # Taking page number (0 index) and splitting page to pages to be extracted 
            pdfWriter.addPage(pdf.getPage(page_num - 1))
        # adding page to pdf file with key as the name
        with open(rf"{cwd}\{k}.pdf", 'wb') as f:
            pdfWriter.write(f)
            f.close

# Function creating new folder and moving LLE Files into folder
def store_files(cwd, page_nums, folder_name):
    dir = folder_name
    parent_dir = cwd
    # Creating full path with file name
    path = os.path.join(parent_dir,dir)
    # Creating directory, its okay if one already exists with that name (kept getting error?)
    os.makedirs(path, exist_ok=True)
    #looping through key values to move each created file from create_files to the new directory
    for key in page_nums.keys():
        shutil.move(rf"{cwd}\{key}.pdf", rf"{cwd}\{dir}\{key}.pdf")

# Function to clear areas on LLE Extractor excel sheet in the most chaotic way possible, im sorry
def clear_sheet():
    extractor_wb = xw.Book('LLE-Extractor-v3.0.xlsm')
    home_sheet = extractor_wb.sheets['Home']
    data_sheet = extractor_wb.sheets['Detailed Data']
    electrical_sheet = extractor_wb.sheets['Electrical Drawing Data']
    mechanical_sheet = extractor_wb.sheets['Mechanical Drawing Data']
    plumbing_sheet = extractor_wb.sheets['Plumbing Drawing Data']

    col_names = {4:"D",5:"E",6:"F",7:"G",8:"H",9:"I",10:"J",11:"K",12:"L",13:"M",14:"N",15:"O",16:"P",17:"Q",18:"R",19:"S",20:"T",21:"U",22:"V",23:"W",24:"X",25:"Y",26:"Z",27:"AA",28:"AB",29:"AC",30:"AD",31:"AE",32:"AF",33:"AG"}
    for num, col in col_names.items():
        data_sheet.range(f'{col}4:{col}20').clear_contents()
    column_also = {13:"M",14:"N",15:"O",16:"P",17:"Q",18:"R",19:"S",20:"T",21:"U",22:"V",23:"W",24:"X",25:"Y",26:"Z",27:"AA"}
    for num, col in column_also.items():
        home_sheet.range(f'{col}5:{col}50').clear_contents()
    data_sheet.range('A4:A500').clear_contents()
    data_sheet.range('B4:B500').clear_contents()
    home_sheet.range('M5:M500').clear_contents()
    home_sheet.range('B7').merge_area.clear_contents()
    home_sheet.range('B10').merge_area.clear_contents()
    home_sheet.range('B13').merge_area.clear_contents()
    home_sheet.range('C15').clear_contents()
    home_sheet.range('F15').clear_contents()
    electrical_sheet.range('A4:A500').clear_contents()
    electrical_sheet.range('B4:B500').clear_contents()
    mechanical_sheet.range('A4:A500').clear_contents()
    mechanical_sheet.range('B4:B500').clear_contents()
    plumbing_sheet.range('A4:A500').clear_contents()
    plumbing_sheet.range('B4:B500').clear_contents()

# Main Script Logic -------------------------------------------------------
def run_LLE():

    # Defining current directory where program is running from
    cwd = r"Z:\Global\0.0 - Administrative - Management\Document Control\Tools\LLE Extraction Tool"
    # Get user inputs - Excel sheet with drawing numbers, drawing set, new file name.
    # Assigning required variables for functions to sheet names, document names, cell ranges, etc
    #Opening Extractor Workbook
    extractor_wb = xw.Book(rf'{cwd}\LLE-Extractor-v3.0.xlsm')
    #Viewing worksheets available
    #worksheet = xw.sheets
    # Defining sheets in extractor workbook
    home_sheet = extractor_wb.sheets['Home']
    data_sheet = extractor_wb.sheets['Detailed Data']
    electrical_sheet = extractor_wb.sheets['Electrical Drawing Data']
    mechanical_sheet = extractor_wb.sheets['Mechanical Drawing Data']
    plumbing_sheet = extractor_wb.sheets['Plumbing Drawing Data']
    #Reading user inputs (Excel sheet, PDF drawings, and name for the output file) and assigning them variable names 
    register = home_sheet.range('B7').value
    register = rf"{cwd}\{register}"
    drawing_set = home_sheet.range('B10').value
    drawing_set = rf"{cwd}\{drawing_set}"
    folder_name = home_sheet.range('B13').value
    package_type = home_sheet.range('C15').value
    stories = home_sheet.range('F15').value
    # Opening Design Deliverables Workbook
    register_wb = xw.Book(register)
    # Defining sheets in design deliverables workbook
    pac_sheet = register_wb.sheets['PAC-DWG']
   # Defining minimum and maximum row numbers for searching for x values, by package type
    if stories == "1s":
        m_row_min = 26
        m_row_max = 58
        p_row_min = 60
        p_row_max = 83
        e_row_min = 85
        e_row_max = 109
    elif stories == "2s":
        m_row_min = 26
        m_row_max = 61
        p_row_min = 63
        p_row_max = 94
        e_row_min = 96
        e_row_max = 126

    # Calling functions to gather page data. Determining how to gather page data using conditional statements and assigning variable to returned dictionary
    if package_type == "Electrical":
        # Calling function to collect electrical data from excel sheet
        page_data = get_drawing_data(electrical_sheet)
    elif package_type == "Mechanical":
        page_data = get_drawing_data(mechanical_sheet)
    elif package_type == "Plumbing":
        # Calling function to get data from drawing set, assigning dictionary to variable name page_data
        page_data = get_drawing_data(plumbing_sheet)
        
    # Calling function to print all extracted page details into Detailed Data sheet for review
    print_detailed_data(data_sheet, page_data)

    #Calling function to return dictionary with key value being package category (row 20 in PAC sheet) and value being list of required drawings
    package_req = get_package_req(pac_sheet, package_type, m_row_min, m_row_max, p_row_min,  p_row_max, e_row_min, e_row_max)

    #Calling function to print required drawings to excel
    print_req_drawings(data_sheet, package_req)
    
    # Calling function to gather list of required page numbers to extract for LLE Packages and print missing drawings to sheet
    page_nums = get_req_page_num(home_sheet, page_data, package_req)

    create_files(cwd, drawing_set, page_nums)

    store_files(cwd, page_nums, folder_name)

# Function only gathering and displaying data so user can fix errors without building a new LLE package each time
def data_only():

    # Defining current directory where program is running from
    cwd = r"Z:\Global\0.0 - Administrative - Management\Document Control\Tools\LLE Extraction Tool"
    # Get user inputs - Excel sheet with drawing numbers, drawing set, new file name.
    # Assigning required variables for functions to sheet names, document names, cell ranges, etc
    #Opening Extractor Workbook
    extractor_wb = xw.Book(rf'{cwd}\LLE-Extractor-v3.0.xlsm')
    #Viewing worksheets available
    #worksheet = xw.sheets
    # Defining sheets in extractor workbook
    home_sheet = extractor_wb.sheets['Home']
    data_sheet = extractor_wb.sheets['Detailed Data']
    electrical_sheet = extractor_wb.sheets['Electrical Drawing Data']
    mechanical_sheet = extractor_wb.sheets['Mechanical Drawing Data']
    plumbing_sheet = extractor_wb.sheets['Plumbing Drawing Data']
   #Reading user inputs (Excel sheet, PDF drawings, and name for the output file) and assigning them variable names 
    register = home_sheet.range('B7').value
    register = rf"{cwd}\{register}"
    drawing_set = home_sheet.range('B10').value
    drawing_set = rf"{cwd}\{drawing_set}"
    package_type = home_sheet.range('C15').value
    stories = home_sheet.range('F15').value
    # Opening Design Deliverables Workbook
    # Opening Design Deliverables Workbook
    register_wb = xw.Book(register)
    # Defining sheets in design deliverables workbook
    pac_sheet = register_wb.sheets['PAC-DWG']
    # Defining minimum and maximum row numbers for searching for x values, by package type
    if stories == "1s":
        m_row_min = 26
        m_row_max = 58
        p_row_min = 60
        p_row_max = 83
        e_row_min = 85
        e_row_max = 109
    elif stories == "2s":
        m_row_min = 26
        m_row_max = 61
        p_row_min = 63
        p_row_max = 94
        e_row_min = 96
        e_row_max = 126

   # Calling functions to gather page data. Determining how to gather page data using conditional statements and assigning variable to returned dictionary
    if package_type == "Electrical":
        # Calling function to collect electrical data from excel sheet
        page_data = get_drawing_data(electrical_sheet)
    elif package_type == "Mechanical":
        page_data = get_drawing_data(mechanical_sheet)
    elif package_type == "Plumbing":
        # Calling function to get data from drawing set, assigning dictionary to variable name page_data
        page_data = get_drawing_data(plumbing_sheet)
        
    # Calling function to print all extracted page details into Detailed Data sheet for review
    print_detailed_data(data_sheet, page_data)

    #Calling function to return dictionary with key value being package category (row 20 in PAC sheet) and value being list of required drawings
    package_req = get_package_req(pac_sheet, package_type, m_row_min, m_row_max, p_row_min,  p_row_max, e_row_min, e_row_max)

    #Calling function to print required drawings to excel
    print_req_drawings(data_sheet, package_req)
    
    # Calling function to gather list of required page numbers to extract for LLE Packages and print missing drawings to sheet
    get_req_page_num(home_sheet, page_data, package_req)