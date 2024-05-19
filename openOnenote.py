import pyautogui
from pywinauto.application import Application
import time

# import osRelated as oR

class oO_class:

    def __init__(self):
        self.app = Application(backend="uia").start(r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE")
        self.app = Application(backend="uia").connect(title = self.get_window_title(), timeout = 6) #try putting self inside the brackets
 
    def get_window_title(self):
        top_window = self.app.top_window()
        window_title = top_window.texts()[0] if top_window.texts() else None
        # print(window_title)
        return window_title
    
    def avoid_special_characters(self, withSpeciaCharacter):
        special_characters = set(r'., *:;#@!$%^()_+={}[]|\:;"<>?\'/')
        withoutSpecialCharacter = ''.join(char for char in withSpeciaCharacter if char not in special_characters)
        return withoutSpecialCharacter

    def nameOfTheNoteBook(self):
        notebookName = input("What is your name of the Note book : ")
        # print(notebookName)
        return notebookName

    def givingNameToNote(self):
        notebookNameGiven = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title="", control_type="Edit").wrapper_object()
        bookName = self.nameOfTheNoteBook()
        # time.sleep(1)
        notebookNameGiven.type_keys(' '+bookName,with_spaces = True) # not getting the first letter of the word due to some unknown reason (so we are giving a extra space! fix it later)

    def openNewNotebook(self):

        fileMenu = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title="File Tab", auto_id="FileTabButton", control_type="Button").wrapper_object()
        fileMenu.click_input()

        newMenu = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title="New", control_type="ListItem").wrapper_object()
        newMenu.click_input()

        # app.Lec22OneNote.print_control_identifiers()

        self.givingNameToNote()

        creatingNotebook = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title="Create Notebook", control_type="Button").wrapper_object()
        creatingNotebook.click_input()

        # alreadyExistingName = app[avoid_special_characters(get_window_title())].child_window(title="OK", auto_id="2", control_type="Button").wrapper_object()
        notNow = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title="Not now", auto_id="2", control_type="Button").wrapper_object() #giving not now as default (to invite people)

        # if alreadyExistingName.texts():
        #     alreadyExistingName.click_input()
        #     givingNameToNote()
            
        # else:
        #     notNow.click_input()

        notNow.click_input()

        #trial
        top_window = self.app.top_window()
        top_window.print_control_identifiers()

    def addPageButton(self):
        addPage = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title="Add Page", control_type="Button").wrapper_object()
        addPage.click_input()

    def printPdf(self):
        insert = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title="Insert", auto_id="TabInsert", control_type="TabItem").wrapper_object()
        insert.click_input()
        filePrintout = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title="File Printout", auto_id="PrintoutFromFileInsert", control_type="Button").wrapper_object()
        filePrintout.click_input()

        address = self.app[self.avoid_special_characters(self.get_window_title())].child_window(title=r"Address: C:\Users\DELL\Downloads", auto_id="1001", control_type="ToolBar").wrapper_object()
        rectangle = address.rectangle()
        new_x_coordinate = rectangle.right - 35

        pyautogui.click(new_x_coordinate, rectangle.mid_point()[1])

        pyautogui.hotkey('ctrl', 'v')

        #trial
        top_window = self.app.top_window()
        top_window.print_control_identifiers()

        
        # address.click_input(coords=(new_x_coordinate, rectangle.mid_point()[1]))
        # # address.click_input(coords=(rectangle.mid_point()[0], rectangle.mid_point()[1]))
        # print((new_x_coordinate, rectangle.mid_point()[1]))

        # app[avoid_special_characters(get_window_title())].type_keys('^v')
    
oO_obj = oO_class()
# oO_obj.printPdf()