# import tkinter as tk
# from tkinter import ttk, messagebox
import os
import json
from pywinauto.application import Application
from pywinauto import Desktop
import psutil
import pandas as pd
from time import sleep
import warnings
from pywinauto.keyboard import send_keys, KeySequenceError
import pywinauto.application
from tkinter import Toplevel, Label
import time
import sys
import win32com.client as com


# def extract_text_from_grid(control):
#     texts = []
    
#     # Check if the control is of type MSHFlexGridWndClass
#     if "Pane" in control._control_types:
#         # Get the window text
#         text = control.window_text()
#         if text:
#             texts.append(text)
    
#     # Recursively check all child controls
#     for child in control.children():
#         texts.extend(extract_text_from_grid(child))
    
#     return texts



df = pd.read_excel('C:/Users/user/OneDrive/Desktop/Bin_Saba/test_file.xlsx')


warnings.simplefilter('ignore', category=UserWarning)

# Pywinauto part 
pids = psutil.pids()
atmgrid = 0
for pid in pids:
    p = psutil.Process(pid)
    if p.name().find("SB_Group.exe") == 0:
        atmgrid = pid
        print(atmgrid)






app = pywinauto.application.Application(backend="win32").connect(process=atmgrid)
dlg = app.window(best_match='SB_Group 247', top_level_only=False)

dlg.set_focus()
for part_num in df["Part Number"]:
    # print(part_num)
    data  = ''
    dlg.set_focus()
    # app.dlg.print_control_identifiers()
    search_box = dlg.Edit
    search_box.set_focus()
    search_box.type_keys(part_num + '{ENTER}', with_spaces=True)

    # flexgrid_control = dlg['MSHFlexGridWndClass13']
    Edit_control = dlg['Edit16']
    
    Edit_control.set_focus()
    # Edit_control.click()
    # Send "Home" key to navigate to the first row
    # Edit_control.type_keys("+{DOWN}")
    data = Edit_control.window_text()
    new_data = ''
    # flexgrid_control = dlg.child_window(class_name="MSHFlexGridWndClass", instance=13)
    flexgrid_control = dlg['      ArabicMSHFlexGridWndClass5']
    flexgrid_control.set_focus()
    
    flexgrid_control.click_input(coords=(15, 18))
    flexgrid_control.type_keys("{ENTER}")
    time.sleep(1)

    while True:
        time.sleep(5)
        Edit_control.set_focus()
        data = Edit_control.window_text()
    
        # Break if the data hasn't changed (indicating we might be at the end or stuck)
        if data == new_data:
            break

        print(data)
        new_data = data

        # Move to the next row
        flexgrid_control.type_keys("{DOWN}{ENTER}")
        

    
    # dlg.print_control_identifiers()
    # all_texts = extract_text_from_grid(dlg)

    # dlg.Pane13.set_focus()
    # dlg.Pane13.click_input()
    # dlg.Pane13.type_keys('{TAB}{ENTER}')


    # grid = com.GetObject(Class="MSEditLib.MSFlexGrid")
    #window = app.window(handle=0x30048)
    #print(window.window_text())

    # for child in dlg.Pane11.children():
    #     print(child.window_text())


    # dlg.set_focus()
    # sys.stdout = open("control_identifiers.txt", "w")
    # dlg.print_control_identifiers()
    # sys.stdout.close()
    # break
# dlg.set_focus()
# sys.stdout = open("control_identifiers.txt", "w")
# dlg.print_control_identifiers()
# sys.stdout.close()


# def get_texts_from_specific_controls(control):
#     texts = []
    
#     # Check the control type using element_info.control_type
#     control_type = control.element_info.control_type
    
#     # Check if the control is of type "Text", "Static", or "Group"
#     if control.is_visible() and control.window_text() and control_type in [u"Text", u"Static", u"Group"]:
#         texts.append(control.window_text())
    
#     # Recursively check all child controls
#     for child in control.childcren():
#         texts.extend(get_texts_from_specific_controls(child))
    
#     return texts


# all_texts = get_texts_from_specific_controls(dlg)
# for text in all_texts:
#     print(text)









# edit.type_keys('', with_spaces=True)
# edit.type_keys(file_name + '{ENTER}' , with_spaces=True)


# try:
#     edit.Save.click()
# except:
#     app.dlg.YesButton.click_input()

# # Uploading with Drive API
# uploading_script(credentials, file_name)

