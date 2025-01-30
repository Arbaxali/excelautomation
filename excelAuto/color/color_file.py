from openpyxl import load_workbook, Workbook
from openpyxl.utils.cell import column_index_from_string
# from openpyxl.utils import cell
import xlwings as xw
import time

def copy_ColorData_to_Temp(exfileReadResolutionData, excel_file_temp):

    try:
        wb1 = xw.Book(exfileReadResolutionData)
        wxb1s = wb1.sheets[0]

        # Open workbook 2

        # alexis, richard, noface

        print("Workbook is being loaded")

        Alexiscolor = []

        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([wxb1s.range((row, 3 + i)).value for row in range(2, 8)])

        for i, row in enumerate(range(3, 9)):
            Alexiscolor.append([wxb1s.range((row, 3 + i)).value for row in range(20, 26)])
        #CW 20
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])
        #cw 80
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #hdr 
        for i, row in enumerate(range(3, 9)):
            Alexiscolor.append([wxb1s.range((row, 3 + i)).value for row in range(44, 50)])

        # A80
        for i, row in enumerate(range(3, 9)):
            Alexiscolor.append([wxb1s.range((row, 3 + i)).value for row in range(38, 44)])

        #d50 80
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #d65 80
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #A_fg
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #d65_fg
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])
        #ww 80
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #cw 250
        for i, row in enumerate(range(3, 9)):
            Alexiscolor.append([wxb1s.range((row, 3 + i)).value for row in range(74, 80)])

        #A 250
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #d50 250
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #d65 250
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #A_fg 250
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #d65_fg 250
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #ww 250
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #d50 350
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #d65 350
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])

        #d65 500
        for i, row in enumerate(range(3, 9)):
            Alexiscolor.append([wxb1s.range((row, 3 + i)).value for row in range(92, 98)])

        #d50 500
        for i, row in enumerate(range(3,9)):
            Alexiscolor.append([0,0,0,0,0,0])




        NofaceColor = []

        for i, row in enumerate(range(3, 9)):
            NofaceColor.append([wxb1s.range((row, 3 + i)).value for row in range(8, 14)])

        for i, row in enumerate(range(3, 9)):
            NofaceColor.append([wxb1s.range((row, 3 + i)).value for row in range(26, 32)])

        #CW 20
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])
        #cw 80
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        for i, row in enumerate(range(3, 9)):
            NofaceColor.append([wxb1s.range((row, 3 + i)).value for row in range(56, 62)])

        for i, row in enumerate(range(3, 9)):
            NofaceColor.append([wxb1s.range((row, 3 + i)).value for row in range(50, 56)])

        #d50 80
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #d65 80
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #A_fg
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #d65_fg
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])
        #ww 80
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        for i, row in enumerate(range(3, 9)):
            NofaceColor.append([wxb1s.range((row, 3 + i)).value for row in range(80, 86)])

        #A 250
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #d50 250
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #d65 250
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #A_fg 250
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #d65_fg 250
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #ww 250
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #d50 350
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        #d65 350
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])


        for i, row in enumerate(range(3, 9)):
            NofaceColor.append([wxb1s.range((row, 3 + i)).value for row in range(98, 104)])

        #d50 500
        for i, row in enumerate(range(3,9)):
            NofaceColor.append([0,0,0,0,0,0])

        RichardColor = []

        for i, row in enumerate(range(3, 9)):
            RichardColor.append([wxb1s.range((row, 3 + i)).value for row in range(14, 20)])

        for i, row in enumerate(range(3, 9)):
            RichardColor.append([wxb1s.range((row, 3 + i)).value for row in range(32, 38)])
        #CW 20
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])
        #cw 80
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        for i, row in enumerate(range(3, 9)):
            RichardColor.append([wxb1s.range((row, 3 + i)).value for row in range(68, 74)])

        for i, row in enumerate(range(3, 9)):
            RichardColor.append([wxb1s.range((row, 3 + i)).value for row in range(62, 68)])

        #d50 80
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #d65 80
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #A_fg
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #d65_fg
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])
        #ww 80
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])


        for i, row in enumerate(range(3, 9)):
            RichardColor.append([wxb1s.range((row, 3 + i)).value for row in range(86, 92)])


        #A 250
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #d50 250
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #d65 250
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #A_fg 250
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #d65_fg 250
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #ww 250
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #d50 350
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        #d65 350
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        for i, row in enumerate(range(3, 9)):
            RichardColor.append([wxb1s.range((row, 3 + i)).value for row in range(104, 110)])

        #d50 500
        for i, row in enumerate(range(3,9)):
            RichardColor.append([0,0,0,0,0,0])

        wb1.close()
        time.sleep(1)


        wxbtemp = xw.Book(excel_file_temp)
        wxb1s = wxbtemp.sheets['Richard']
        wxb2s = wxbtemp.sheets['Alexis']
        wxb3s = wxbtemp.sheets['Noface']
        ColumnStart = 68
        RichardRowStart = 232
        NofaceRowStart = 232
        AlexiRowStart = 232



        for rowData in Alexiscolor:                
            wxb2s[chr(ColumnStart) + str(AlexiRowStart)].value = rowData
            AlexiRowStart +=1
        print("Alexis color data is been copied to temp")


        for rowData in NofaceColor:                
            wxb3s[chr(ColumnStart) + str(NofaceRowStart)].value = rowData
            NofaceRowStart +=1
        print("Noface color data is been copied to temp")



        for rowData in RichardColor:                
            wxb1s[chr(ColumnStart) + str(RichardRowStart)].value = rowData
            RichardRowStart +=1
        print("Richaard color data is been copied to temp")


        wxbtemp.save()
        wxbtemp.close()


    except Exception as e:

        print("An error occurred while reading and updaing temp excel:", e)
        error_message  = str(e)
        start_index = error_message.find("'") + 1
        end_index = error_message.rfind(".xlsx'") + len(".xlsx")
        extracted_filename = error_message[start_index:end_index]
        print(extracted_filename)
