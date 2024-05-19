import osRelated as oR
import openOnenote as oO

# class main_class(oO.oO_class, oR.oR_class):
    

# oR_obj = oR.oR_class()
# oO_obj = oO.oO_class()

oR.oR_obj.unzip_folders(oR.oR_obj.input_folder, oR.oR_obj.output_folder)
oO.oO_obj.openNewNotebook()
oR.oR_obj.folder_1by1(oR.oR_obj.output_folder)

#trial
oO.oO_obj.printPdf()

# main_obj = main_class()