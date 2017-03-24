'''
Created on Mar 7, 2017

@author: logans
'''
import openpyxl as xl

def main():
    wb_original = xl.load_workbook('FDE_List.xlsx') # This file contains the old, poorly-formatted FDEs
    wb_new = xl.Workbook() # This will be the new excel file for receiving the cleaned up strings
    fde_sheet = wb_original.active # should return the only workbook in the document
    list_of_clean_strings = []
    
    for row in fde_sheet.rows:
        for cell in row:
            #if not cell.value: continue # If the cell's contents are empty, skip the mofo
            dirty_string = str(cell.value)
            dirty_string = " ".join(dirty_string.split())
            while ' ,' in dirty_string:
                dirty_string = dirty_string.replace(' ,', ',')
            
            while ', ' in dirty_string:
                dirty_string = dirty_string.replace(', ', ',')
                
            list_of_clean_strings.append(dirty_string)
    
    print("finished appending all clean items to temporary list") # debug statement
    print("Size of clean strings: " + str(len(list_of_clean_strings)))
    
    for i in range(0,len(list_of_clean_strings)):
        list_of_clean_strings[i] = list_of_clean_strings[i].replace(",", ",\r")
    
    print("addition of newline characters complete")
    
    print(list_of_clean_strings[-1])
    
#     for row in list_of_clean_strings:
#         print(row)
   
    sheet_new = wb_new.active #name = "Sheet" by default
    sheet_new.title = "Fixed FDEs"
    
    for i in range(1,len(list_of_clean_strings)+1):
        str_i = "A" + str(i) # loading the cell coordinate into a string 
        sheet_new[str_i] = str(list_of_clean_strings[i-1])
    wb_new.save('Fixed_FDEs.xlsx')
    print("Fixed_FDEs.xlsx created successfully")
    
    #list_of_clean_strings = [] # This is here so Python garbage collects this big list and reclaims the memory
    

    
main()
