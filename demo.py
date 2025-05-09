# -*- coding: utf-8 -*-
# from pywinauto import application, Desktop
from pywinauto import Desktop
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.application import Application, ProcessNotFoundError
from pywinauto.timings import Timings
from pywinauto.controls.uiawrapper import UIAWrapper
from typing import Optional
import psutil
import os

Timings.slow()
Timings.window_find_timeout = 10
Timings.after_click_wait=5
Timings.after_listviewselect_wait=5
Timings.after_comboboxselect_wait=5

key_vs = "Microsoft Visual Studio"
key_vs_dlg =f"{key_vs}Dialog"

def get_vs_executable(dist:str = 'Enterprise') -> str:
    '''Get the Visual Studio executable path based on the distribution.
    Returns:
        str: The path to the Visual Studio executable.
    '''
    default_vs_folder = os.path.join(os.environ.get('PROGRAMFILES'), "Microsoft Visual Studio", "2022")
    vs_dists = ['Community', 'Professional', 'Enterprise']
    if dist not in vs_dists:
        raise ValueError("Invalid distribution specified.")
    return f"{default_vs_folder}\\{dist}\\Common7\\IDE\\devenv.exe"

app_vs_executable = get_vs_executable()
app = Application(backend="uia", allow_magic_lookup=True)

def kill_app_by_process_id(process_id:int):
    '''Kill the application process by its ID.
    Args:
        process_id (int): The ID of the process to kill.
    '''
    try:
        p = psutil.Process(process_id)
        p.terminate()
        p.wait(timeout=3)
    except psutil.NoSuchProcess:
        print(f"Process {process_id} not found.")
    except psutil.AccessDenied:
        print(f"Access denied to terminate process {process_id}.")
    except Exception as e:
        print(f"Error terminating process {process_id}: {e}")

def kill_app_by_name(name:str):
    '''Kill the application process by its name.
    Args:
        name (str): The name of the process to kill.
    '''
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == name:
            try:
                proc.terminate()
                proc.wait(timeout=3)
            except psutil.NoSuchProcess:
                print(f"Process {proc.info['pid']} not found.")
            except psutil.AccessDenied:
                print(f"Access denied to terminate process {proc.info['pid']}.")
            except Exception as e:
                print(f"Error terminating process {proc.info['pid']}: {e}")

def start_app(key:str):
    '''Start the application and wait for it to be ready.
    Args:
        key (str): The key to identify the application window.
    '''
    app.start(app_vs_executable)
    app[key].wait('ready')
    # app[key_vs].print_control_identifiers()

def connect_app(key:str):
    '''Connect to the application and wait for it to be ready.
    Args:
        key (str): The key to identify the application window.
    '''
    try:
        app.connect(path = app_vs_executable)
        app[key].wait('ready')
        return app[key]
    except ProcessNotFoundError:
        return None
    except Exception as e:
        print(f"Error connecting to {key}: {e}")
        raise

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

def find_windows(doc_name:str)-> Optional[UIAWrapper]:
    '''Find a window by its document name.
    Args:
        doc_name (str): The document name of the window to find.
    Returns:
        UIAWrapper: The found window object.
    '''
    try:
        return Desktop(backend="uia").window(title=doc_name)
    except Exception as e:
        print(f"Error finding window {doc_name}: {e}")
        return None

def test_start_app():
    '''Test the start_app function.'''
    start_app(key_vs)
    app[key_vs].wait('ready')
    app[key_vs].print_control_identifiers()

def test_connect_app():
    '''Test the connect_app function.'''
    menu_view_error_list = "View->Error List"
    menu_view_output = "View->Output"
    menu_build_solution = "Build->Build Solution"
    proj_name = "ConsoleApp9"
    key_vs_opened_project = f"{proj_name} - {key_vs}"
    
    app[key_vs].wait('ready')
    # app[key_vs].print_control_identifiers()
    title_bar = app[key_vs_opened_project][key_vs_opened_project]
    win_menu_select(title_bar, menu_view_output)
    win_menu_select(title_bar, menu_view_error_list)

if __name__ == "__main__":
    
    kill_app_by_name("devenv.exe")
    
     
    start_app(key_vs)
    
    a = find_windows(key_vs)
    a.minimize()
    
    a.restore()
    print(type(a))
   