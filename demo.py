# -*- coding: utf-8 -*-
# from pywinauto import application, Desktop
from pywinauto.application import Application, WindowSpecification
from pywinauto.timings import Timings
from pywinauto.controls.uiawrapper import UIAWrapper


Timings.slow()
Timings.window_find_timeout = 10
Timings.after_click_wait=5
Timings.after_listviewselect_wait=5
Timings.after_comboboxselect_wait=5

dist = 'Community'
dist = 'Enterprise'

app_executable = r"C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\devenv.exe"

key_vs = "Microsoft Visual Studio"
key_vs_dlg =f"{key_vs}Dialog"

app = Application(backend="uia", allow_magic_lookup=True)

def start_app(key:str):
    '''Start the application and wait for it to be ready.
    Args:
        key (str): The key to identify the application window.
    '''
    app.start(app_executable)
    app[key].wait('ready')
    # app[key_vs].print_control_identifiers()

def connect_app(key:str):
    '''Connect to the application and wait for it to be ready.
    Args:
        key (str): The key to identify the application window.
    '''
    app.connect(path = app_executable)
    app[key].wait('ready')
    # app[key_vs].print_control_identifiers()

# app.start(app_executable)

# app[key_vs].wait('ready')
# app[key_vs].print_control_identifiers()

# app[key_vs_dlg].child_window(auto_id = 'Button_1', control_type='button', title='').click()

# app[key_vs][key_vs].menu_select("File->New->Project")
# filer project template
# app[key_vs][key_vs_dlg].child_window(title='LanguageFilter').select("C#")
# app[key_vs][key_vs_dlg].child_window(title='PlatformFilter').select("Windows")
# app[key_vs][key_vs_dlg].child_window(title='ProjectTypeFilter').select("Console")
# select Console App template
# app[key_vs][key_vs_dlg].child_window(title='Project Templates', control_type='List', auto_id='ListViewTemplates').item("Console App").select()
# magic match
# app[key_vs][key_vs_dlg].ProjectTemplates.item("Console App").select()
# app[key_vs][key_vs_dlg].child_window(auto_id='button_Next').click()
# configure your new project
# app[key_vs][key_vs_dlg].child_window(auto_id='projectNameText').type_keys("MyConsoleApp")
# proj_name = "ConsoleApp9"
# proj_name = app[key_vs][key_vs_dlg].child_window(auto_id='projectNameText').texts()
# print(proj_name)
# app[key_vs][key_vs_dlg].child_window(auto_id='locationCmb').type_keys(r"C:\temp").type_keys("{TAB}")
# app[key_vs][key_vs_dlg].child_window(auto_id='button_Next').click()
# app[key_vs][key_vs_dlg].child_window(title='_Framework', class_name='ComboBox', auto_id="ComboBoxControl").select(".NET 8.0 (Long Term Support)")
# app[key_vs][key_vs_dlg].child_window(auto_id='button_Next', title='Create').click()
# app[key_vs].wait('ready')
# VisualStudioMainWindow
# key_vs_created_project = f"{proj_name} - {key_vs}"
# app.connect(path = app_executable)
# app[key_vs_created_project].wait('ready')
# app[key_vs_created_project][key_vs_created_project].menu_select("Build->Build Solution")
# app[key_vs_created_project][key_vs_created_project].menu_select("View->Error List")

def click_continue_without_code(key: str):
    '''Click the "Continue without code" button in Visual Studio.
    Args:
        key (str): The key to identify the application window.
    '''
    app[key].child_window(title="Continue without code", control_type="Button").click()

def win_menu_select(wrapper, menu:str):
    '''Select a menu item from the application.
    Args:
        key (WindowSpecification): The WindowSpecification object to identify the application window.
        menu (str): The menu item to select.
    '''
    wrapper.menu_select(menu)

def menu_select(key:str, menu:str):
    '''Select a menu item from the application.
    Args:
        key (str): The key to identify the application window.
        menu (str): The menu item to select.
    '''
    # app[key].child_window(title=key, class_name="Window", auto_id=key).wait('ready')
    win_menu_select(app[key][key], menu)

# Error List
# TabGroup|ST:0:0:{d78612c7-9962-4b83-95d9-268046dad23a}|ST:0:0:{34e76e81-ee4a-11d0-ae2e-00a0c90fffc3}
# Name: Error List
# ST:0:0:{d78612c7-9962-4b83-95d9-268046dad23a}
# Tracking List View
# app[key_vs_created_project].ErrorList.print_control_identifiers()
# app[key_vs_created_project][key_vs_created_project].child_window(title="Error List", class_name="ToolBar").print_control_identifiers()
# app[key_vs_created_project].child_window(title="Results", auto_id="Tracking List View", control_type="Table").print_control_identifiers()

# SolutionExplorer
# 

def get_solution_explorer(key:str):
    # Solution Explorer
    app[key].child_window(title="Solution Explorer", control_type="Tree", auto_id="SolutionExplorer").print_control_identifiers()
    # app[key_vs_created_project].child_window(title='Solution Explorer', class_name='TreeView', auto_id='SolutionExplorer').print_control_identifiers()

if __name__ == "__main__":
    
    menu_view_error_list = "View->Error List"
    menu_view_output = "View->Output"
    menu_build_solution = "Build->Build Solution"
    proj_name = "ConsoleApp9"
    key_vs_opened_project = f"{proj_name} - {key_vs}"
    start_app(key_vs)
    connect_app(key_vs)
    # app[key_vs][key_vs].print_control_identifiers()
    
    title_bar = app[key_vs][key_vs]
    

    click_continue_without_code(key_vs)
    win_menu_select(title_bar, menu_view_output)
    win_menu_select(title_bar, menu_view_error_list)
    
    # app[key_vs_opened_project].child_window(title="MenuBar", auto_id="MenuBar", control_type="MenuBar").menu_select("View->Error List")
    # app[key_vs_opened_project][key_vs_opened_project].menu_select("View->Error List")
    # menu_select(key_vs_opened_project, "View->Error List")
    # get_solution_explorer(key_vs_opened_project)