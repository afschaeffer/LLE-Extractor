#-----------------------------------------------------------------------------------------------------
# 
# Script Name: LLE-Extraction.py
# Author:      Amy Schaeffer
#
# Version: 1
# Start Date:        8/30/2022
# End Date:          9/22/2022
# Description: Identifies required drawing numbers from a range in a user defined Excel 
#              sheet and parses through a Drawing Set PDF, extracting related
#              pages and compiling them into a new PDF file. Additionally, prints list of extracted
#              page data into separate sheet in excel for further review if needed 
# 
#----------------------------------------------------------------------------------------------------

# Importing xlwings to read and print in Excel
from ast import Num
from numbers import Number
import xlwings as xw
# Importing PDFMiner Libraries
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator


# Search for instance of needed drawing numbers in PDF data dictionary, creating new list of key value pairs with data (drawing number from excel file: page number from parsed data)
# If drawing name required is not found in parsed PDF, add drawing name to new list and
#  output missing drawings into extraction tool

# Extract and compile required drawing sheets from each column into new PDF using PyPDF4 

# Save new PDF file to specified path

# Function to add passed key and value to passed dictionary
def add_values_in_dict(sample_dict, key, value):
    sample_dict[key] = value
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
def get_e_drawing_data(electrical_drawing_sheet):
    ret = {}
    # Looping through range between rows 4 and 500 for formatting purpose
    for row in range(4,500):
        # Adding values in columns A and B to dictionary until it hits a blank cell
        cell_value = electrical_drawing_sheet[f'A{row}'].value
        if cell_value != None:
            key = electrical_drawing_sheet[f'A{row}'].value
            value = electrical_drawing_sheet[f'B{row}'].value
            value = format_e_page_num(value)
            ret[key] = value
        else:
            break
    return ret

# Function parsing passed drawing set PDF, gathering drawing number and title from passed related coordinates, incrementing page number and storing and returning data in a dictionary
def get_drawing_set_data(drawing_set, drawing_number_xmin, drawing_number_xmax, drawing_number_ymin, drawing_number_ymax):
    # Opening PDF and creating parser object
    fp = open(drawing_set, 'rb')
    manager = PDFResourceManager()
    laparams = LAParams()
    dev = PDFPageAggregator(manager, laparams=laparams)
    interpreter = PDFPageInterpreter(manager, dev)
    pages = PDFPage.get_pages(fp)
    
    # Creating empty page number variable and page data dict
    page_num = 0
    page_data = {}
    
    #Looping through each page, parsing it, and collecting and storing drawing number and page data
    for page in pages:
        interpreter.process_page(page)
        layout = dev.get_result()

        for lobj in layout: 
            # Checking if LTTextbox is within drawing number coordinates, storing text if true
             if (lobj.bbox[0] >= drawing_number_xmin) and (lobj.bbox[1] >= drawing_number_ymin) and (lobj.bbox[2] <= drawing_number_xmax) and (lobj.bbox[3] <= drawing_number_ymax):
                try:
                    drawing_number = lobj.get_text()
                    #removing additional spaces in front of and at end of string
                    drawing_number = drawing_number.strip()
                    #drawing_number = format_drawing_number(drawing_number)
                except:
                    continue
        # Incrementing page number
        page_num = page_num + 1
        # Adding page values into dictionary, with drawing number as the key and list with drawing title and page number as the value pair
        add_values_in_dict(page_data, drawing_number, page_num)
    # Returning data in a dictionary
    return page_data

# Function to update Data Details sheet with all key value pairs found in page data dictionary
def print_detailed_data(destination_sheet, dictionary):
    #creating counters for printing data into cells in each column
    key_counter = 4
    value_counter = 4
    #clearing contents of columns prior to printing
    destination_sheet.range('A4:A500').clear_contents()
    destination_sheet.range('B4:B500').clear_contents()
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
    col_names = {27:"AA",28:"AB",29:"AC",30:"AD",31:"AE",32:"AF",33:"AG",34:"AH",35:"AI",36:"AJ",37:"AK",38:"AL",   39:"AM",40:"AN",41:"AO",42:"AP",43:"AQ",44:"AR",45:"AS"}

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
    # Dictionary holding column indexes that will need cleared prior to printing to excel
    col_names = {4:"D",5:"E",6:"F",7:"G",8:"H",9:"I",10:"J",11:"K",12:"L",13:"M",14:"N",15:"O",16:"P",17:"Q",18:"R",19:"S",20:"T",21:"U",22:"V",23:"W",24:"X",25:"Y",26:"Z",27:"AA",28:"AB",29:"AC",30:"AD",31:"AE",32:"AF",33:"AG"}
    # Clearing contents of columns prior to printing
    for num, col in col_names.items():
        sheet.range(f'{col}4:{col}20').clear_contents()
    # Printing each key value into column D in sheet, incrementing counter after each iteration
    for key, value in dictionary.items():
        sheet.range(f'D{key_counter}').options(index=False).value = key
        key_counter = key_counter +1
    #printing each value from list into columns E+, incrementing counter after each iteration
    for values, value in dictionary.items():
        sheet.range(f'E{value_counter}').options(index=False).value = value
        value_counter = value_counter +1  

# Main Script Logic -------------------------------------------------------

# PDF Pixel coordinate values for drawing title and drawing number, for passing to get_drawing_set_data method
drawing_number_xmin = 2740
drawing_number_ymin = 70
drawing_number_xmax = 2963
drawing_number_ymax = 96

# Get user inputs - Excel sheet with drawing numbers, drawing set, new file name.
# Assigning required variables for functions to sheet names, document names, cell ranges, etc
#Opening Extractor Workbook
extractor_wb = xw.Book('LLE-Extractor-v1.0.xlsm')
#Viewing worksheets available
worksheet = xw.sheets
# Defining sheets in extractor workbook
home_sheet = extractor_wb.sheets['Home']
data_sheet = extractor_wb.sheets['Detailed Data']
electrical_sheet = extractor_wb.sheets['Electrical Drawing Data']
#Reading user inputs (Excel sheet, PDF drawings, and name for the output file) and assigning them variable names 
design_deliverables = home_sheet.range('B7').value
drawing_set = home_sheet.range('B10').value
new_file_name = home_sheet.range('B13').value
package_type = home_sheet.range('C15').value
# Opening Design Deliverables Workbook
design_deliverables_wb = xw.Book(design_deliverables)
# Defining sheets in design deliverables workbook
pac_sheet = design_deliverables_wb.sheets['PAC']
# Defining minimin and maximum row numbers for searching for x values, by package type
m_row_min = 26
m_row_max = 68
p_row_min = 70
p_row_max = 93
e_row_min = 95
e_row_max = 119

# Calling functions to gather page data. Determining how to gather page data using conditional statements and assigning variable to returned dictionary
if package_type == "Electrical":
    # Calling function to collect electrical data from excel sheet
    page_data = get_e_drawing_data(electrical_sheet)
elif package_type == "Mechanical" or package_type == "Plumbing":
    # Calling function to get data from drawing set, assigning dictionary to variable name page_data
    page_data = get_drawing_set_data(drawing_set, drawing_number_xmin, drawing_number_xmax, drawing_number_ymin, drawing_number_ymax)
    
# Calling function to print all extracted page details into Detailed Data sheet for review
print_detailed_data(data_sheet, page_data)

#Calling function to return dictionary with key value being package category (row 20 in PAC sheet) and value being list of required drawings
package_req = get_package_req(pac_sheet, package_type, m_row_min, m_row_max, p_row_min,  p_row_max, e_row_min, e_row_max)

#Calling function to print required drawings to excel
print_req_drawings(data_sheet, package_req)

