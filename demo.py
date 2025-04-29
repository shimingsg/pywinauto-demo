# -*- coding: utf-8 -*-
from pywinauto import application, timings, Desktop

timings.Timings.slow()
timings.Timings.after_click_wait=5
timings.Timings.after_listviewselect_wait=5
timings.Timings.after_comboboxselect_wait=5

dist = 'Community'
dist = 'Enterprise'

app_executable = r"C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\devenv.exe"

key_vs = "Microsoft Visual Studio"
key_vs_dlg =f"{key_vs}Dialog"

app = application.Application(backend="uia")
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
proj_name = "ConsoleApp9"
# proj_name = app[key_vs][key_vs_dlg].child_window(auto_id='projectNameText').texts()
# print(proj_name)
# app[key_vs][key_vs_dlg].child_window(auto_id='locationCmb').type_keys(r"C:\temp").type_keys("{TAB}")
# app[key_vs][key_vs_dlg].child_window(auto_id='button_Next').click()
# app[key_vs][key_vs_dlg].child_window(title='_Framework', class_name='ComboBox', auto_id="ComboBoxControl").select(".NET 8.0 (Long Term Support)")
# app[key_vs][key_vs_dlg].child_window(auto_id='button_Next', title='Create').click()
# app[key_vs].wait('ready')
# VisualStudioMainWindow
key_vs_created_project = f"{proj_name} - {key_vs}"
app.connect(path = app_executable)
app[key_vs_created_project].wait('ready')
# app[key_vs_created_project][key_vs_created_project].menu_select("Build->Build Solution")
app[key_vs_created_project][key_vs_created_project].menu_select("View->Error List")

# Error List
# TabGroup|ST:0:0:{d78612c7-9962-4b83-95d9-268046dad23a}|ST:0:0:{34e76e81-ee4a-11d0-ae2e-00a0c90fffc3}
# Name: Error List
# ST:0:0:{d78612c7-9962-4b83-95d9-268046dad23a}
# Tracking List View
app[key_vs_created_project][key_vs_created_project].ErrorList.print_control_identifiers()
# app[key_vs_created_project][key_vs_created_project].child_window(title="Error List", class_name="ToolBar").print_control_identifiers()

# Show items contained by
# 
