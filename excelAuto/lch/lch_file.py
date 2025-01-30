from openpyxl import load_workbook, Workbook
from openpyxl.utils.cell import column_index_from_string
# from openpyxl.utils import cell
import xlwings as xw
import time


def copy_from_LCHData_to_Temp(exfileReadResolutionData,excel_file_temp):
    try:
        wxb1 = xw.Book(exfileReadResolutionData)
        wxb1s = wxb1.sheets['Sheet1']
        AlexisResolution = []
        print("Workbook loaded")
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([wxb1s.range((row, 5 + i)).value for row in range(2, 8)])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([wxb1s.range((row, 5 + i)).value for row in range(20, 26)])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([wxb1s.range((row, 5 + i)).value for row in range(26, 32)])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([wxb1s.range((row, 5 + i)).value for row in range(32, 38)])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([wxb1s.range((row, 5 + i)).value for row in range(8, 14)])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([wxb1s.range((row, 5 + i)).value for row in range(14, 20)])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             AlexisResolution.append([0,0,0,0,0,0])

        NofaceResolution = []
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([wxb1s.range((row, 5 + i)).value for row in range(38, 42)])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([wxb1s.range((row, 5 + i)).value for row in range(56, 62)])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([wxb1s.range((row, 5 + i)).value for row in range(62, 68)])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([wxb1s.range((row, 5 + i)).value for row in range(68, 74)])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([wxb1s.range((row, 5 + i)).value for row in range(44, 50)])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([wxb1s.range((row, 5 + i)).value for row in range(50, 56)])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             NofaceResolution.append([0,0,0,0,0,0])

        RichardResolution = []
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([wxb1s.range((row, 5 + i)).value for row in range(74, 80)])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([wxb1s.range((row, 5 + i)).value for row in range(92, 98)])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([wxb1s.range((row, 5 + i)).value for row in range(98, 104)])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([wxb1s.range((row, 5 + i)).value for row in range(104, 110)])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([wxb1s.range((row, 5 + i)).value for row in range(80, 86)])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([wxb1s.range((row, 5 + i)).value for row in range(86, 92)])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])
        for i, row in enumerate(range(5, 7)):
             RichardResolution.append([0,0,0,0,0,0])

        wxb1.close()
        time.sleep(1)

        wxbtemp = xw.Book(excel_file_temp)
        wxb1s = wxbtemp.sheets['Richard']
        wxb2s = wxbtemp.sheets['Alexis']
        wxb3s = wxbtemp.sheets['Noface']
        ColumnStart = 68
        RichardRowStart = 114
        NofaceRowStart = 114
        AlexiRowStart = 114

        for rowData in AlexisResolution:                 
                wxb2s[chr(ColumnStart) + str(AlexiRowStart)].value = rowData
                AlexiRowStart +=1
        print("Alexis lch data being copied to temp")

        for rowData in NofaceResolution:                
                wxb3s[chr(ColumnStart) + str(NofaceRowStart)].value = rowData
                NofaceRowStart +=1
        print("noface lch data being copied to temp")

        for rowData in RichardResolution:                
                wxb1s[chr(ColumnStart) + str(RichardRowStart)].value = rowData
                RichardRowStart +=1
        print("Richard lch data being copied to temp")

        wxbtemp.save()
        wxbtemp.close()

        print("Lch data copied to temp")

    except Exception as e:
        print("An error occurred while reading and updaing temp excel:", e)
        error_message  = str(e)
        start_index = error_message.find("'") + 1
        end_index = error_message.rfind(".xlsx'") + len(".xlsx")
        extracted_filename = error_message[start_index:end_index]
        print(extracted_filename)
        # closing workbooks if opened with same names
