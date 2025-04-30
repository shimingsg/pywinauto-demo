
# ComboBox

``` python
app[key_vs][key_vs_dlg].child_window(control_type="ComboBox", found_index=0).select("F#")
app[key_vs][key_vs_dlg].child_window(auto_id="ComboBox_1").select("C#")
```

# ListView
``` python
app[key_vs][key_vs_dlg].child_window(auto_id='ListViewTemplates', title='Project Templates', control_type='List').item("Console App").select()
app[key_vs][key_vs_dlg].child_window(auto_id='ListViewTemplates', title='Project Templates', control_type='List').item(1).select()
```

``` python
from pywinauto.application import Application, ProcessNotFoundError
from pywinauto import findwindows

timeout = 20

def start_application(app_executable):
    try:
        app = Application(backend="uia").connect(path=app_executable)
    except ProcessNotFoundError:
        print(f"Process not found: {app_executable}")
        app = Application(backend="uia").start(app_executable, timeout=timeout)
    
    dlg = app.window(title="Microsoft Visual Studio")
    print("Process started")
    return dlg
```
