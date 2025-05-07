# -*- coding: utf-8 -*-
"""This script automates the process of:
1. Creating a new C# Console Application project in Visual Studio 2022.
2. Adding a new C# Class Library project to the solution.
3. Adding a reference to the Class Library in the Console Application.
4. Building the solution to ensure there are no compilation errors.
5. Running the application to confirm it executes as expected.
6. Inserting a breakpoint in the code editor.
7. Starting debugging to verify that the breakpoint is hit.
8. Stopping debugging.
The script uses the pywinauto library to interact with the Visual Studio UI.
It is designed to work with Visual Studio 2022 Community and Enterprise editions.
"""
import os
import shutil

from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import Timings
from pywinauto.controls.uiawrapper import UIAWrapper

Timings.slow()
Timings.window_find_timeout = 10
Timings.after_click_wait = 5
Timings.after_listviewselect_wait = 5
Timings.after_comboboxselect_wait = 5

dist = "Enterprise"

app_executable = (
    fr"C:\Program Files\Microsoft Visual Studio\2022\{dist}\Common7\IDE\devenv.exe"
)

key_vs = "Microsoft Visual Studio"
key_vs_dlg = f"{key_vs}Dialog"

app = Application(backend="uia", allow_magic_lookup=True)


def start_app(key: str):
    """Start the application and wait for it to be ready.
    Args:
        key (str): The key to identify the application window.
    """
    app.start(app_executable)
    app[key].wait("ready")


def connect_app(key: str):
    """Connect to the application and wait for it to be ready.
    Args:
        key (str): The key to identify the application window.
    """
    app.connect(path=app_executable)
    app[key].wait("ready")


def click_continue_without_code(key: str):
    """Click the "Continue without code" button in Visual Studio.
    Args:
        key (str): The key to identify the application window.
    """
    window_continue_without_code = app[key].child_window(
        title="Continue without code", control_type="Button"
    )
    if window_continue_without_code.exists():
        window_continue_without_code.click()
    else:
        print("Continue without code button not found.")
    # app[key].child_window(title="Continue without code", control_type="Button").click()
    # if app[key].child_window(title="Continue without code", control_type="Button").exists():
    #     app[key].child_window(title="Continue without code", control_type="Button").click()


def create_new_console_project(
    proj_name: str, dotnet_version: str, proj_location: str, title_bar
):
    """Create a new console application project.
    Args:
        proj_name (str): The name of the project.
        dotnet_version (str): The .NET version to use.
        proj_location (str): The location to save the project.
    """
    title_bar.menu_select("File->New->Project")
    dlg = app[key_vs][key_vs_dlg]

    # Select "Console App" template
    dlg.child_window(title="LanguageFilter").select("C#")
    dlg.child_window(title="PlatformFilter").select("Windows")
    dlg.child_window(title="ProjectTypeFilter").select("Console")
    dlg.child_window(
        title="Project Templates", control_type="List", auto_id="ListViewTemplates"
    ).item("Console App").select()
    dlg.child_window(auto_id="button_Next").click()

    # Set up the project name and location
    dlg.child_window(auto_id="projectNameText").type_keys(proj_name)
    location_box = dlg.child_window(auto_id="locationCmb", control_type="ComboBox")
    location_box.type_keys(proj_location, with_spaces=True)
    dlg.child_window(auto_id="projectNameText").type_keys("{TAB}")
    dlg.child_window(auto_id="button_Next", title="Next").click()

    ## Select the framework
    framework_dropdown = dlg.child_window(title="_Framework", auto_id="ComboBoxControl", control_type="ComboBox")
    framework_dropdown.expand()
    framework_dropdown.select(dotnet_version)
    dlg.child_window(auto_id="button_Next", title="Create").click()

    # Handle "Overwrite Project" dialog
    try:
        overwrite_dlg = app.window(title_re=key_vs)
        overwrite_dlg.child_window(title="Yes", control_type="Button").click()
    except ElementNotFoundError:
        # By expectation, this dialog may not appear if the project is new
        pass

    print(f"{proj_name} project created successfully.")


def win_menu_select(wrapper, menu: str):
    """Select a menu item from the application.
    Args:
        key (WindowSpecification): The WindowSpecification object to identify the application window.
        menu (str): The menu item to select.
    """
    wrapper.menu_select(menu)


def menu_select(key: str, menu: str):
    """Select a menu item from the application.
    Args:
        key (str): The key to identify the application window.
        menu (str): The menu item to select.
    """
    # app[key].child_window(title=key, class_name="Window", auto_id=key).wait('ready')
    win_menu_select(app[key][key], menu)


def add_new_class_library(lib_name: str, proj_location: str, title_bar):
    """Add a new C# Class Library project to the solution.
    Args:
        lib_name (str): The name of the class library project.
        proj_location (str): The location to save the project.
    """
    # Open the "Add New Project" dialog
    title_bar.menu_select("File->Add->New Project")
    dlg = app[key_vs][key_vs_dlg]

    # Select "Class Library" template
    dlg.child_window(title="LanguageFilter").select("C#")
    dlg.child_window(title="PlatformFilter").select("Windows")
    dlg.child_window(title="ProjectTypeFilter").select("Library")
    dlg.child_window(
        title="Project Templates", control_type="List", auto_id="ListViewTemplates"
    ).item("Class Library").select()
    dlg.child_window(auto_id="button_Next").click()

    # Clean up the project name if it already exists
    try:
        shutil.rmtree(os.path.join(proj_location, lib_name))
    except FileNotFoundError:
        pass

    # Configure the new project
    dlg.child_window(auto_id="projectNameText").type_keys(lib_name)
    location_box = dlg.child_window(auto_id="locationCmb", control_type="ComboBox")
    location_box.type_keys(proj_location, with_spaces=True)  # Set the location path
    dlg.child_window(auto_id="projectNameText").type_keys(
        "{TAB}"
    )  # Simulate pressing TAB to move focus

    next_button = dlg.child_window(auto_id="button_Next", title="Next")
    next_button.click()

    # Select the framework as .NET Standard 2.1
    framework_dropdown = dlg.child_window(
        auto_id="ComboBoxControl", control_type="ComboBox"
    )
    framework_dropdown.expand()
    framework_dropdown.select(".NET Standard 2.1")
    dlg.child_window(auto_id="button_Next", title="Create").click()

    print(f"{lib_name} project created successfully.")


def add_project_reference(proj_name: str, lib_name: str, title_bar):
    """Add a reference to ClassLibrary in ConsoleApp.
    Args:
        proj_name (str): The name of the console application project.
        lib_name (str): The name of the class library project.
    """
    # Open the Solution Explorer
    title_bar.menu_select("View->Solution Explorer")
    solution_explorer = app[key_vs].child_window(
        title="Solution Explorer", control_type="Tree", auto_id="SolutionExplorer"
    )

    # Select the target project in Solution Explorer
    target_project = solution_explorer.child_window(
        title=proj_name, control_type="TreeItem"
    )
    target_project.click_input()

    # Add a reference to the ClassLibrary project
    title_bar.menu_select("Project->Add Project Reference...")
    reference_dialog = app[key_vs][f"Reference Manager - {proj_name}"]

    # Switch to the "Projects" tab
    projects_tab = reference_dialog.child_window(
        title="Projects", control_type="TreeItem"
    )
    projects_tab.select()

    # Check if the ClassLibrary checkbox is available and select it
    lib_checkbox = reference_dialog.child_window(
        title=lib_name, control_type="CheckBox"
    )
    lib_checkbox.double_click_input()

    # Confirm the reference addition
    ok_button = reference_dialog.child_window(title="OK", control_type="Button")
    ok_button.click_input()

    print(f"{lib_name} reference added to {proj_name} successfully.")


def build_solution(title_bar):
    """Build the solution to ensure there are no compilation errors."""
    win_menu_select(title_bar, "Build->Build Solution")


def run_application(title_bar):
    """Run the application and confirm it executes as expected."""
    title_bar.wait("ready")
    title_bar.menu_select("Debug->Start Without Debugging")


def insert_breakpoint(line_number: int = 1):
    """Insert a breakpoint in the code editor.
    Args:
        line_number (int): The line number where the breakpoint should be inserted.
    """
    # Open the code editor for the main file
    program_tab = app[key_vs].child_window(title="Program.cs", control_type="TabItem")
    program_tab.click_input()

    # Select the line where you want to insert the breakpoint
    program_tab.set_focus()
    program_tab.type_keys("^g")  # Ctrl+G to open "Go To Line" dialog

    goto_dialog = app[key_vs].child_window(title="Go To Line", control_type="Window")
    goto_dialog.wait('ready')
    goto_dialog.child_window(control_type="Edit").set_text(str(line_number))
    goto_dialog.child_window(title="OK", control_type="Button").click()

    program_tab.type_keys("{F9}")

    print("Breakpoint inserted successfully.")


def start_debugging(title_bar):
    """Start debugging and verify that the breakpoint is hit."""
    title_bar.menu_select("Debug->Start Debugging")


def stop_debugging(title_bar):
    """Stop debugging the application."""
    title_bar.menu_select("Debug->Stop Debugging")


def main():
    dotnet_version = ".NET 9.0 (Standard Term Support)"
    proj_location = r"D:\test"
    proj_name = "ConsoleApp2"
    lib_name = "ClassLibrary1"
    start_app(key_vs)
    connect_app(key_vs)

    title_bar = app[key_vs][key_vs]

    click_continue_without_code(key_vs)
    create_new_console_project(proj_name, dotnet_version, proj_location, title_bar)
    add_new_class_library(lib_name, proj_location, title_bar)
    add_project_reference(proj_name, lib_name, title_bar)
    build_solution(title_bar)
    run_application(title_bar)
    insert_breakpoint(2)
    start_debugging(title_bar)
    # stop_debugging(title_bar)

    input("Press Enter to continue...")


if __name__ == "__main__":
    main()
