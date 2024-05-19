import os
import shutil
import pyperclip

# import openOnenote_fns as oO
import openOnenote as oO
# import main

class oR_class:

    # def __init__(self):# is it mandatory to give self
        # self.input_folder = r'F:\projects\zipToOneNoteAutomation\trial_folder'
        # self.output_folder = r'F:\projects\zipToOneNoteAutomation\trial_folder\unZipped_folders'

    # class variable
    input_folder = r'F:\projects\zipToOneNoteAutomation\trial_folder'
    output_folder = r'F:\projects\zipToOneNoteAutomation\trial_folder\unZipped_folders'

    # using the staticmethod inside the class (since it's not affecting any instance variable particualarly)
    @staticmethod
    def get_pdf_files(directory): #for getting .pdf in all levels inside a folder #not going to use!
        pdf_files = []
        for root, dirs, files in os.walk(directory): #root give the name of the current dir ; dirs are the dir inside the current dir ; files are the files inside the current dir.
            for file in files:
                if file.endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))
        return pdf_files

    @staticmethod
    def unzip_folders(input_path, output_path):
        zip_files = [file for file in os.listdir(input_path) if file.endswith('.zip')] #will this line go inside other folders to find .zip file
        for zip_file in zip_files:
            zip_file_path = os.path.join(input_path, zip_file)
            shutil.unpack_archive(zip_file_path, output_path)

    # Process to to run through all unzipped folders
    @staticmethod
    def folder_1by1(output_path):
        entries = os.listdir(output_path)
        subdirectories_folder = [entry for entry in entries if os.path.isdir(os.path.join(output_path, entry))]
        for i in subdirectories_folder:
            # oO.oO_class.addPageButton()
            oO.oO_obj.addPageButton()
            print(i)
        
    # Process to print all .pdf file in different pages in a given section
    @staticmethod
    def pdf_1by1(dir):
        entries = os.listdir(dir)
        # subdirectories_pdf = [entry for entry in entries if entry.lower().endswith('.pdf') and os.path.isfile(os.path.join(dir, entry))]
        # print(subdirectories_pdf)

        # flag = 0
        pdf_dir_list = [os.path.join(dir, entry) for entry in entries if entry.lower().endswith('.pdf') and os.path.isfile(os.path.join(dir, entry))]
        # print(pdf_dir_list)
        pyperclip.copy(pdf_dir_list)

    # input_folder = r'F:\projects\zipToOneNoteAutomation\trial_folder'
    # output_folder = r'F:\projects\zipToOneNoteAutomation\trial_folder\unZipped_folders'

    # Unzip folders
    # unzip_folders(input_folder, output_folder) #By default it's not unzipping the zipped the file if there already exist the unzipped version whose content is same(not even when same content with different zip file name)

    # # Process each PDF file
    # pdf_files = get_pdf_files(output_folder)
    # for pdf_file in pdf_files:
    #     # add_to_onenote(pdf_file)
    #     print('hello')

    # folder_1by1(output_folder)

# it's better to make the obj of class in the module/file itself otherwise it cause circular import
oR_obj = oR_class()

# currently in unzip_folders() there is no any exception handling system if we are unzipping a file of the same name to the 'unZipped_folders' of the existing files.