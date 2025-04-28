# -*- coding: utf-8 -*-
from pywinauto import application
from pywinauto import timings

dist = 'Community'
dist = 'Enterprise'

app_executable = r"C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe"

key_vs = "Microsoft Visual Studio"
key_vs_dlg ="Microsoft Visual StudioDialog"
app = application.Application(backend="uia")
app.start(app_executable)

app[key_vs].wait('ready')
# app[key_vs].print_control_identifiers()

app[key_vs_dlg].child_window(auto_id = 'Button_1', control_type='button', title='').click()

# app[key_vs][key_vs].menu_select("File->New->Project")
timings.Timings.slow()
timings.Timings.after_click_wait=5
timings.Timings.after_listviewselect_wait=5
timings.Timings.after_comboboxselect_wait=5
# filer project template
# app[key_vs][key_vs_dlg].child_window(title='LanguageFilter').select("C#")
# app[key_vs][key_vs_dlg].child_window(title='PlatformFilter').select("Windows")
# app[key_vs][key_vs_dlg].child_window(title='ProjectTypeFilter').select("Console")
# select Console App template
app[key_vs][key_vs_dlg].child_window(title='Project Templates', control_type='List', auto_id='ListViewTemplates').item("Console App").select()
# magic match
# app[key_vs][key_vs_dlg].ProjectTemplates.item("Console App").select()
app[key_vs][key_vs_dlg].child_window(auto_id='button_Next').click()
# configure your new project
app[key_vs][key_vs_dlg].child_window(auto_id='projectNameText').type_keys("MyConsoleApp")
app[key_vs][key_vs_dlg].child_window(auto_id='locationCmb').type_keys(r"C:\temp")



# app[key_vs][key_vs_dlg].ProjectTemplates.item("Console App").select()
# .select("Console App", found_index=0)
# app[key_vs][key_vs_dlg].child_window(control_type="Edit", found_index=0).type_keys("MyConsoleApp")
