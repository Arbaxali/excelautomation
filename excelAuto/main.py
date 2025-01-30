from challenge.alexis_challenge import read_AppliedFormulaDatat_to_alexis_Add_to_Temp  
#from challenge.alexis_challengetest import read_AppliedFormulaDatat_to_alexis_Add_to_Temp
from challenge.noface_challenge import read_AppliedFormulaDatat_to_noface_Add_to_Temp
from challenge.richard_challenge import read_AppliedFormulaDatat_to_richard_Add_to_Temp
from formula.apply_formula import apply_Formula_to_Run
from ae.ae_file import copy_from_ae_file
from awb.awb_file import copy_from_awb_file
from ht.ht_file import copy_from_ht_file
from resolution.resolution_file import copy_ResolutionData_to_Temp
from color.color_file import copy_ColorData_to_Temp
from lch.lch_file import copy_from_LCHData_to_Temp
import time
import os


if __name__ == "__main__":

    start = time.time()
    # directory_path = 'C:\\Users\\yadavdix\\Music\\Input_Log\\'
    directory_path = "C:\\Users\\arbazalx\\Downloads\\Madhavi-HF-Golden-Set1\\"
    xlsx_files = [file for file in os.listdir(directory_path) if file.endswith(".xlsx")]
    Alexis_list = []
    Richard_list = []
    Noface_list = []

    for file_name in xlsx_files:
        if "Alexis" in file_name:
            Alexis_list.append(directory_path + file_name)
        elif "Richard" in file_name:
            Richard_list.append(directory_path + file_name)
        elif "Noface" in file_name:
            Noface_list.append(directory_path + file_name)
        elif "color" in file_name:
            colorexfile = directory_path + file_name
        elif "AWB" in file_name:
            awbexfile = directory_path + file_name
        elif "AE" in file_name:
            aeexfile = directory_path + file_name
        elif "resolution" in file_name:
            resolutionexfile = directory_path + file_name
        elif "LCH" in file_name:
            lchfaceexfile = directory_path + file_name
        elif "HT" in file_name:
            htexfile = directory_path + file_name
        elif "temp" in file_name:
            tempexfile = directory_path + file_name

# # richard data
    for _excel in Richard_list:
        apply_Formula_to_Run(_excel)

    for _excel in Richard_list:
        read_AppliedFormulaDatat_to_richard_Add_to_Temp(_excel,tempexfile)
    

# alexis data

    for _excel in Alexis_list:
           apply_Formula_to_Run(_excel)

    for _excel in Alexis_list:
        read_AppliedFormulaDatat_to_alexis_Add_to_Temp(_excel,tempexfile) 


# no face data 
    for _excel in Noface_list:
           apply_Formula_to_Run(_excel)

    for _excel in Noface_list:
           read_AppliedFormulaDatat_to_noface_Add_to_Temp(_excel,tempexfile)



    copy_from_ae_file(aeexfile, tempexfile)
    copy_from_awb_file(awbexfile,tempexfile)
    copy_from_ht_file(htexfile, tempexfile)

    copy_ResolutionData_to_Temp(resolutionexfile, tempexfile)
    copy_ColorData_to_Temp(colorexfile, tempexfile)
    copy_from_LCHData_to_Temp(lchfaceexfile,tempexfile)

    end = time.time()
    print("operation completed")
    print("The time of execution of above program is :", (end-start) , "s")