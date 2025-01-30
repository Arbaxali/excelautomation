from openpyxl import load_workbook, Workbook
from openpyxl.utils.cell import column_index_from_string
# from openpyxl.utils import cell
import xlwings as xw
import time



def copy_from_ht_file(exhtfile, excel_file_temp):
    try:
        print("Loading ht workbook")      
        wxb1 = xw.Book(exhtfile)
        wxb1s = wxb1.sheets['Sheet1']      
        ht_readAlexisDataContainer = []  
        ht_readRichirdDataContainer = []  
        n = 0
        Alexisranges = [
            (16, 26), (28, 38), (40, 50), (52, 62), (64, 74)
        ]

        Rechardranges = [
            (76, 86), (88, 98), (100, 110), (112, 122), (124, 134)
        ]

        for range_start, range_end in Alexisranges:
            ht_readAlexisDataContainer.append([wxb1s.range((row, 2)).value for row in range(range_start, range_end+1, 2)])
            ht_readAlexisDataContainer.append([wxb1s.range((row, 3)).value for row in range(range_start, range_end+1, 2)]) 

        for range_start, range_end in Rechardranges:
            ht_readRichirdDataContainer.append([wxb1s.range((row, 2)).value for row in range(range_start, range_end+1, 2)])
            ht_readRichirdDataContainer.append([wxb1s.range((row, 3)).value for row in range(range_start, range_end+1, 2)])

 
        wxb1.close()

        wxbtemp = xw.Book(excel_file_temp)
        wxb1s = wxbtemp.sheets['Richard']
        wxb2s = wxbtemp.sheets['Alexis']
        ColumnStart = 68
        RichardRowStart = 180
        AlexiRowStart = 180
 
        for rowData in ht_readRichirdDataContainer:                
                wxb1s[chr(ColumnStart) + str(RichardRowStart)].value = rowData
                RichardRowStart +=1
 

        for rowData in ht_readAlexisDataContainer:                
                wxb2s[chr(ColumnStart) + str(AlexiRowStart)].value = rowData
                AlexiRowStart +=1

        wxbtemp.save()
        wxbtemp.close()
        print("Data from ht is copied")

 

    except Exception as e:

        print("An error occurred:", e)
        error_message  = str(e)
        start_index = error_message.find("'") + 1
        end_index = error_message.rfind(".xlsx'") + len(".xlsx")
        extracted_filename = error_message[start_index:end_index]
        print(extracted_filename)

