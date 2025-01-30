from openpyxl import load_workbook, Workbook
from openpyxl.utils.cell import column_index_from_string
# from openpyxl.utils import cell
import xlwings as xw
import time

def copy_ResolutionData_to_Temp(exfileReadResolutionData,excel_file_temp):

    try:
        wxb1 = xw.Book(exfileReadResolutionData)
        wxb1s = wxb1.sheets['Sheet1']
        AlexisResolution = []
        print("Resolution workbook loaded")

        for i, row in enumerate(range(3, 11)):
             AlexisResolution.append([wxb1s.range((row, 3 + i)).value for row in range(8, 14)])

        for i, row in enumerate(range(3, 11)):
             AlexisResolution.append([wxb1s.range((row, 3 + i)).value for row in range(14, 20)])

        for i, row in enumerate(range(3, 11)):
             AlexisResolution.append([wxb1s.range((row, 3 + i)).value for row in range(2, 8)])


        NofaceResolution = []

        for i, row in enumerate(range(3, 11)):
             NofaceResolution.append([wxb1s.range((row, 3 + i)).value for row in range(26, 32)])

        for i, row in enumerate(range(3, 11)):
             NofaceResolution.append([wxb1s.range((row, 3 + i)).value for row in range(32, 38)])

        for i, row in enumerate(range(3, 11)):
             NofaceResolution.append([wxb1s.range((row, 3 + i)).value for row in range(20, 26)])


        RichardResolution = []

        for i, row in enumerate(range(3, 11)):
             RichardResolution.append([wxb1s.range((row, 3 + i)).value for row in range(44, 50)])

        for i, row in enumerate(range(3, 11)):
             RichardResolution.append([wxb1s.range((row, 3 + i)).value for row in range(50, 56)])
        for i, row in enumerate(range(3, 11)):
             RichardResolution.append([wxb1s.range((row, 3 + i)).value for row in range(38, 44)])

 
        wxb1.close()
        time.sleep(1)


        wxbtemp = xw.Book(excel_file_temp)
        wxb1s = wxbtemp.sheets['Richard']
        wxb2s = wxbtemp.sheets['Alexis']
        wxb3s = wxbtemp.sheets['Noface']
        ColumnStart = 68
        RichardRowStart = 364
        NofaceRowStart = 364
        AlexiRowStart = 364

 

        for rowData in AlexisResolution:                
            wxb2s[chr(ColumnStart) + str(AlexiRowStart)].value = rowData
            AlexiRowStart +=1

        print("Alexis resolution data being copied to temp")

        for rowData in NofaceResolution:                
            wxb3s[chr(ColumnStart) + str(NofaceRowStart)].value = rowData
            NofaceRowStart +=1

        print("noface resolution data being copied to temp")
 

        for rowData in RichardResolution:                
            wxb1s[chr(ColumnStart) + str(RichardRowStart)].value = rowData
            RichardRowStart +=1
        
        print("Richard resolution data being copied to temp")

        wxbtemp.save()
        wxbtemp.close()

 

    except Exception as e:

        print("An error occurred while reading and updaing temp excel:", e)
        error_message  = str(e)
        start_index = error_message.find("'") + 1
        end_index = error_message.rfind(".xlsx'") + len(".xlsx")
        extracted_filename = error_message[start_index:end_index]
        print(extracted_filename)

