# A guide to Windows hotkey automatisation in Python

<img src="./logo.ico" width="128" height="12" align="right" />

## Introduction

This guide aims to show how to build fully customisable hotkeys with Python. They can be used to easily automate repetitive tasks to increase productivity when working on a Windows machine. The use of Python as opposed to specialised programming languages such as AutoHotkey makes the initial setup more complex, but allows complete freedom in the implementation of the hotkey listener, proper multi-threading, a higher reliability and offers far more ready-made solutions, since Python is a much more popular. Originally, the guide is designed for Windows 10, but all examples have also been tested on Windows 7 and 8.

## Build a Python hotkey listener

### Package PyWinhook

#### Installation

The [PyWinhook](https://github.com/Tungsteno74/pyWinhook) package enables us to interact with the Windows user input processing directly in Python. This allows triggering Python routines by pressing a hotkey combination. It also ensures that the keyboard input sent to the Python script is not additionally processed by another currently running program.

We can install the package using:`pip install pyWinhook`

Although it is not listed as a requirement, experience has shown that the installation fails if Swig is not installed on the system. It can be downloaded [here](https://sourceforge.net/projects/swig) and must be added to PATH.

#### Event loop

Here is a basic setup for a hotkey listener with PyWinhook:
```python
import os  # used for script termination
import pythoncom  # used for event loop
import pyWinhook  # used as keyboard listener
import threading  # used for multi-threading

def Hotkey():  # Routines
    print('Exit')
    os._exit(0)

def KeyPress(event):  # Listener
    if event.Key == 'Escape':
        thread = threading.Thread(target=Hotkey)
        thread.start()
        return False
    else:
        return True

hook = pyWinhook.HookManager()  # Initalisation
hook.KeyDown = KeyPress
hook.HookKeyboard()
pythoncom.PumpMessages()
```
The programme is divided into three sections, the initialisation, the listener and the routines. First, the initialisation is executed, which creates a keyboard input listener and then passes all events to it. The latter is implemented with the command`pythoncom.PumpMessages()`, which acts similar to a while true loop, so all lines below it are never executed.

Each time a keyboard key is pressed, the listener is called as the`KeyPress`function and the object event is passed. For the beginning, however, we are only interested in the attribute`.Key`, which contains the current key as a string. If the key pressed was *Esc* (`'Escape'`), false is returned, which ensures that this *Esc* is not sent to any other programme. In all other cases, true is returned and the keyboard input is forwarded normally to all other programmes. This demonstrates how important it is to be careful when tinkering with the logic of the listening section, as a wrong implementation could block all keyboard input. Therefore, it is highly advisable to always include an exit condition in the script.

Instead of calling the`Hotkey()`function in the routines section directly, we create a thread for this function and let it run. This has the distinct advantage that the listener does not have to wait until the execution of the function is complete. Otherwise, in a more complex system, the execution of the listener could take several seconds, resulting in significant keyboard input lag. Therefore, the top priority is to make the execution time of the listener as short as possible.

#### Modifier

With the previous programme design, we can only react to individual keyboard keys; calling up a function with a combination such as *Ctrl + U* is not possible. To change this, we now register for some keys that we see as modifiers (e.g. *Ctrl*) when exactly they are pressed.

```python
import os
import pythoncom
import pyWinhook
import threading

def Hotkey():  # Ctrl & Esc
    print('Exit')
    os._exit(0)

def KeyPress(event):
    if event.Key == 'Lshift' or event.Key == 'Rshift':  # Shift
        modifier[0] = True
        return True
    elif event.Key == 'Lcontrol' or event.Key == 'Rcontrol':  # Ctrl
        modifier[1] = True
        return True
    elif event.Key == 'Lmenu' or event.Key == 'Rmenu':  # Alt
        modifier[2] = True
        return True
    elif event.Key == 'Escape':
        if modifier[1]:
            thread = threading.Thread(target=Hotkey)
            thread.start()
            return False
        else:
            return True
    else:
        return True

def KeyRelease(event):
    if event.Key == 'Lshift' or event.Key == 'Rshift':  # Shift
        modifier[0] = False
        return True
    elif event.Key == 'Lcontrol' or event.Key == 'Rcontrol':  # Ctrl
        modifier[1] = False
        return True
    elif event.Key == 'Lmenu' or event.Key == 'Rmenu':  # Alt
        modifier[2] = False
        return True
    else:
        return True

modifier = [False, False, False]  # [Shift, Control, Alt]

hook = pyWinhook.HookManager()
hook.KeyDown = KeyPress
hook.KeyUp = KeyRelease
hook.HookKeyboard()
pythoncom.PumpMessages()
```
Since we now also have a function in the listener section that is called as soon as a key is released again, we can store in an array which of the modifier keys are currently pressed. So now the`Hotkey()`function is only called when *Ctrl & Esc* are pressed at the same time. Of course, the selection of modifiers can be customised at will, for example by adding the 10 Numpad keys.

#### Mouse listener

PyWinhook also allows to process the input of the mouse:
```python
import os
import pythoncom
import pyWinhook
import threading

def Hotkey():  # MouseL
    print('Exit')
    os._exit(0)

def KeyPress(event):
    ...

def KeyRelease(event):
    ...

def MousePress(event):
    if event.MessageName == 'mouse left down':
        thread = threading.Thread(target=Hotkey)
        thread.start()
        return False
    else:
        return True

hook = pyWinhook.HookManager()
hook.KeyDown = KeyPress
hook.KeyUp = KeyRelease
hook.HookKeyboard()
hook.SubscribeMouseAllButtonsDown(MousePress)
hook.HookMouse()
pythoncom.PumpMessages()
```

Here the event object is structured differently, the name of the button is stored in`.MessageName`, while`.Position`contains the current position on the screen.

### Autostart setup

#### VBScript

Now that the basics are explained, we need to make sure that the script runs invisibly in the background and is automatically available at every restart.

In order to avoid that the command shell of the programme is distracting in the task bar, we use VBScript. Assuming the script is saved with the name *Automatisation.py* in the folder *C:\Automatisation*, create a file called *Start.vbs* with the following content:
```vbscript
Dim WShell
Set WShell = CreateObject("WScript.Shell")
WShell.run "cmd.exe /c cd C:\Automatisation && python Automatisation.py", 0
Set WShell = Nothing
```
When *Start.vbs* is executed, the programme is active without being externally noticeable. Before executing a script in this way, one should therefore make sure to have implemented an exit condition; otherwise, the script can only be terminated by restarting the system. If you now put the *Start.vbs* in the autostart folder (*%AppData%\Microsoft\Windows\Start Menu\Programs\Startup*), the script is started automatically with every restart.

#### Elevated Privileges

However, this method does not work if one of the functions in the Routines section requires elevated privileges for its work (for example, work in the registry) and so the script must be run as admin. Windows only runs the programmes in the autostart folder with normal permissions, so the only option left is the Task Scheduler. There we can simply create a task, set the action to run *Start.vbs* and select *Run with highest privileges*. The Task Scheduler has even implemented functionality that allows tasks to be run directly at start-up. However, experience has shown that these are very unreliable. To work around this, create another file *Task.vbs* in the autostart folder with the following content (assuming the task created has the name *Automatisation*):
```vbscript
Dim WShell
Set WShell = CreateObject("WScript.Shell")
WShell.run "cmd.exe /c C:\WINDOWS\system32\schtasks.exe /run /i /tn ""Automatisation""", 0
Set WShell = Nothing
```
So at startup, Windows runs *Task.vbs* with normal privileges, which calls the task *Automation* in the Task Scheduler, which in turn calls *Start.vbs* with elevated privileges, so that the Python script can run with elevated privileges.

#### Notification

Generally, if a script is running invisibly in the background, it is advisable to include a hotkey routine to check whether the script is running at all. One solution would be to send a Windows notification, as the following function does (`pip install plyer`necessary):
```python
import plyer

def Notification():
    plyer.notification.notify(
        title = 'Automatisation',
        message = 'Programme is active.',
        timeout = 5,
    )
```

## Programmes, folders and links

### Package Os

To open folders, programmes and links in the Routines Section, we use the Os package. For example, if we want to open the folder *C:\Automatisation*, we implement this function in the routine section:
```python
def OpenScriptPath():
    os.startfile('C:\\Automatisation')
```
Or to open the recycling bin:
```python
def OpenRecyclingBin():
    os.system('C:\\Windows\\SysWOW64\\explorer.exe ::{645FF040-5081-101B-9F08-00AA002F954E}')
```
Similarly we open a link with this function:
```python
def OpenLink():
    os.startfile('https://github.com/PhoenixSmaug/python-automatisation')
```
And we can start the notepad with the following syntax (extra quotes prevent errors from paths with spaces):
```python
def OpenNotepad():
    os.system('"C:\\Windows\\System32\\notepad.exe"')
```
Or open the battery usage statistics in the settings:
```python
def OpenBatteryUsage():
    os.system('C:\\Windows\\SysWOW64\\explorer.exe ms-settings:batterysaver-usagedetails')
```
Again, we get problems when the script runs as administrator. Then all programmes called by `os.system()`run with elevated privileges, which is a security risk and sometimes leads to unwanted behaviour. We can prevent that by implementing the following (assuming the user name is *abc*):
```python
def OpenNotepadNonAdmin():
    os.popen('runas /user:abc /savecred "C:\\Windows\\System32\\notepad.exe"')
```
However, before this is implemented, the command`runas /user:abc /savecred cmd.exe`should be executed once manually in the console. After you have entered your user password once, it will be saved by the`/savecred`flag. This way the Python script can use the command without having to enter the credentials itself.

An alternative way would be to create a separate task in the task scheduler for each program and then disable *Highest privileges*. However, this way no parameters can be passed to the program, which is necessary in some cases (see Functions for extended capabilites). For simple programs it is nevertheless a solution (assuming the task is called *Notepad* in the task folder *Auto*):
```python
def OpenNotepadNonAdmin():
    os.popen('schtasks /run /i /tn ""Auto\\Notepad""')
```

### Strings

Now we will look at how the Python programme itself can input text. For example, you could build a hotkey that automatically fills in your own mail address to shorten logins. The best Python package for this functionality that also supports special characters is keyboard and is installed with`pip install keyboard`. The function in the Routines Section is then structured as follows:
```python
def WriteMail():
    keyboard.write('example@mailprovider.com')
```
With multi-line strings, it is advisable to create the line break with *Shift & Enter*, as many chat apps send the message with a simple *Enter*. This is most easily implemented with the package pyautogui(`pip install pyautogui`) and would look like this:
```python
def WriteGrettings():
    keyboard.write('Kindest regards,')
    pyautogui.hotkey('shift', 'enter')
    keyboard.write('John Doe')
```
However, there is a danger when inserting strings: The keyboard input that the programme generates with`keyboard.write()`is itself passed back to`KeyPress(event)`. So if in an implementation the keyboard press of *e* leads to the output of *hello*, the routine triggers itself and it leads to a crash. Fortunately, the Windows API allows you to distinguish between real user input and keyboard input generated by programmes. So you can avoid the problem by implementing the following return condition in`KeyPress(event)`and`KeyRelease(event)`:
```python
if event.Injected == 16:
    return True
```

### Functions for extended capabilities

#### WindowExist

Now we will talk about auxiliary functions that can make our hotkeys smarter. First, let us look at a function that can be used to check whether a programme is open somewhere on the computer.
```python
import subprocess

def WindowExist(string_window):
    call = 'TASKLIST', '/FI', 'imagename eq %s' % string_window
    output = subprocess.check_output(call).decode()
    last_line = output.strip().split('\r\n')[-1]
    return last_line.lower().startswith(string_window.lower())
```
This implementation uses subprocess, a vanilla Python package and was developed by [ewerbody](https://stackoverflow.com/a/29275361). It can be used like this:
```python
if not WindowExist('firefox.exe'):
    os.startfile('https://weboas.is')
os.startfile('https://github.com/PhoenixSmaug/python-automatisation')
```
This causes the default home tab to be opened first, if Firefox is not already open, before the actual link is opened in a new tab.

#### WindowActive

The next auxiliary function, WindowActive returns the currently focused programme. Later when we will deal with hotkeys specific to the File Explorer, this function allows us to check in the routines whether we are in the Explorer at all.
```python
import psutil
import time
import win32gui
import win32process

def WindowActive():
    pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
    exe = psutil.Process(pid[-1]).name()
    title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    return exe, title
```
The installation of psutil (`pip install psutil`) is required, win32gui and win32process are already present as dependencies of pywinhook. The usage then looks like this:
```python
[ActiveExe, _] = WindowActive()
if ActiveExe == 'explorer.exe':
    pyautogui.hotkey('F1')
```

#### ExplorerPath

Another indispensable functionality is a function to obtain the current path of an open Windows Explorer instance. This allows you to open, for example, VS code in the current path at the touch of a button. Windows does not make such a query very easy, but with inspiration from a post of [DADi590](https://stackoverflow.com/a/52959617) I managed to build a relieable solution:
```python
import urllib.parse
import win32.win32gui
import win32com.client

def ExplorerPath():
    shell = win32com.client.Dispatch('{9BA05972-F6A8-11CF-A442-00A0C90A8F39}')
    for win in range(shell.Count):
        if shell[win].hwnd == win32.win32gui.GetForegroundWindow():
            url = urllib.parse.unquote(shell[win].LocationURL,encoding='ISO 8859-1')
            directory = url.split("///")[1].replace("/", "\\")
            return directory.communicate()[0].decode("utf-8").rstrip()
    return None
```
Urllib is a vanilla package, while the others are already installed by PyWinhook. The applications of the function would then look like this:
```python
def VSCode():  # Open Visual Studio Code in current path
    path = ExplorerPath()
    if path != 'None':
        os.popen('runas /user:abc /savecred "C:\\Users\\abc\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe \\"' + path + '\\""')

def CMD():  # Open console in current path
    path = ExplorerPath()
    if path != 'None':
        subprocess.call('start C:\\Windows\\System32\\cmd.exe', cwd = path, shell = True)

def PowerShell():  # Open powershell in current path
    path = ExplorerPath()
    if path != 'None':
        subprocess.call('start C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe', cwd = path, shell = True)
```

### Better initialisation

#### Permission check

Now we look at various features for the initialisation section. For example, if the script needs to run with elevated privileges for some processes, we should check at startup if it actually has these privileges:
```python
import ctypes
import os

if ctypes.windll.shell32.IsUserAnAdmin() == 0:
    os._exit(0)
```
Using only vanilla packages the script checks whether it has administrator rights and terminates itself otherwise.

#### Modifier state

If *Caps* or the Numpad keys are used as modifiers, it is advisable to ensure that *CapsLock* and *NumLock* are deactivated when starting the script:
```python
import pyautogui
import win32.win32api as win32api

if win32api.GetKeyState(win32con.VK_CAPITAL) == 1:  # CapsLock Off
    pyautogui.typewrite(['capslock'])

if win32api.GetKeyState(win32con.VK_NUMLOCK) == 1:  # NumLock Off
    pyautogui.typewrite(['numlock'])
```
Win32api is already installed as a dependency of PyWinhook.

#### Package schedule

Often, certain routines should not only be triggered by hotkeys, but should also be called automatically every hour (for example, a backup script). The package schedule (`pip install schedule`) can be used for this. Since it runs in an endless loop similar to`pythoncom.PumpMessages()`, threads must be used to run both programmes side by side.
```
import os
import pythoncom
import pyWinhook
import schedule
import threading

def BackupYouTube():
    ...

def BackupFiles():
    ...  

def KeyPress(event):
    ...

def KeyRelease(event):
    ...

def Schedule():
    schedule.every().hour.do(BackupYouTube)  # Examples
    schedule.every().day.at("12:00").do(BackupFiles)

    while True:
        schedule.run_pending()
        time.sleep(1)

thread = threading.Thread(target=Schedule)
thread.start()

hook = pyWinhook.HookManager()
hook.KeyDown = KeyPress
hook.KeyUp = KeyRelease
hook.HookKeyboard()
pythoncom.PumpMessages()
```

## Advanced routines

### Custom Windows controls

#### System Exit

Here I will list a few very practical routines for which the implementation was rather challenging. First, routines to exit the system:
```python
import os
import win32.lib.win32con as win32con
import win32.win32gui as win32gui

def ScreenOff():
    win32gui.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con.SC_MONITORPOWER, 2)

def SleepMode():
    for i in range(len(array_modifier)):  # Mark modifiers as released
        modifier[i] = False
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def Shutdown():
    os.system("shutdown -s -t 0")
```
Before we put the system into sleep mode, it is important to mark all modifiers as released. Otherwise, the programme could miss the actual release of the modifier and would still see the key as pressed after waking up.

#### Window Management

Often, for example, one would like to have a certain window permanently in the foreground in order to be able to transcribe data without having to tediously position the windows next to each other.
```python
import win32.lib.win32con as win32con
import win32.win32gui as win32gui

def WindowTop():
    win32gui.SetWindowPos(win32gui.GetForegroundWindow(), win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

def WindowUntop():
    win32gui.SetWindowPos(win32gui.GetForegroundWindow(), win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
```
This allows you to toggle the Foreground status of the currently focused programme via hotkey.

#### Pixel RGB Value

Another functionality that has proven extremely useful is to be able to access the current RGB values of the pixel under the mouse with a hotkey.

```python
import win32.win32api as win32api
import win32.win32gui as win32gui
import win32.win32clipboard as win32clipboard

def PixelRGB():
    x, y = win32api.GetCursorPos()
    colour = hex(win32gui.GetPixel(win32gui.GetDC(win32gui.GetActiveWindow()), x, y))[2:]
    win32clipboard.OpenClipboard()
    win32clipboard.SetClipboardText(colour, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()
```
It writes the RGB values as hex numbers into the clipboard, again only dependencies of PyWinhook are used.

### Explorer improvements

#### Show hidden files

Now we will focus on applications in the Windows File Explorer. We start with a method that toggles whether the Explorer displays hidden files. Windows does not make it easy, but with a little work in the registry it can be done.
```python
import win32.lib.win32con as win32con
import win32.win32gui as win32gui
import winreg

def ExplorerHide():
    [ActiveExe, _] = WindowActive()
    if ActiveExe == 'explorer.exe':
        handles = []
        RegRead = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", 0, winreg.KEY_READ)
        value, type = winreg.QueryValueEx(RegRead, "Hidden")
        winreg.CloseKey(RegRead)

        RegWrite = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", 0, winreg.KEY_WRITE)
        if value == 2:
            winreg.SetValueEx(RegWrite, "Hidden", 0, winreg.REG_DWORD, 1)
        else:
            winreg.SetValueEx(RegWrite, "Hidden", 0, winreg.REG_DWORD, 2)
        winreg.CloseKey(RegWrite)

        win32gui.EnumWindows(ExplorerFind, None)
        list(map(ExplorerRefresh, handles))
    else:
        ...

def ExplorerFind(index, _):
    global handles
    if win32gui.GetClassName(index) == 'CabinetWClass':
        handles.append(index)

def ExplorerRefresh(index):
    win32gui.PostMessage(index, win32con.WM_COMMAND, 28931, None)
    win32gui.PostMessage(index, win32con.WM_COMMAND, 41504, None)
```
The code first checks whether an Explorer instance is currently focused and if so, it toggles the show-hidden option via the registry. Again, only dependencies of PyWinhook are used. To ensure that the open Explorer windows also react to the new registry value, they must all be refreshed with the auxiliary functions ExplorerFind and ExplorerRefresh developed by [viilpe](https://stackoverflow.com/a/64974981).

#### Zip and unzip

Another frequently-used application is the fast zipping and unzipping of multiple files and folders.
```python
import os
import pyautogui
import time
import win32.win32clipboard as win32clipboard
import zipfile

def Zip():
    [ActiveExe, _] = WindowActive()
    if ActiveExe == 'explorer.exe':
        pyautogui.hotkey('ctrl', 'c')  # Copy selected files to clipboard
        time.sleep(0.1)
        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
            win32clipboard.CloseClipboard()
        except TypeError:  # No files selected
            win32clipboard.CloseClipboard()
            return

        pathExplorer = ExplorerPath()
        archive = zipfile.ZipFile(os.path.join(pathExplorer, os.path.splitext(data[0])[0]) + '.zip', 'w')  # Create ZIP
        for fileSelect in data:
            archive.write(fileSelect, os.path.basename(fileSelect))
            for root, folder, file in os.walk(fileSelect):  # Include all files and folders in directories
                for element in file:
                    archive.write(os.path.join(root, element), os.path.join(root, element).replace(pathExplorer + '\\', ''))
                for element in folder:
                    archive.write(os.path.join(root, element), os.path.join(root, element).replace(pathExplorer + '\\', ''))
        archive.close()
    else:
        ...

def Unzip():
    [ActiveExe, _] = WindowActive()
    if ActiveExe == 'explorer.exe':
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
            win32clipboard.CloseClipboard()
        except TypeError:  # No files selected
            win32clipboard.CloseClipboard()
            return

        for file in data:
            if file[len(file) - 4:] == '.zip':
                zipfile.ZipFile(file,'r').extractall(file[:len(file) - 4])
    else:
        ...
```
With the help of the Vanilla Package zipfile, all selected files can be saved in a zip archive via`Zip()`. And`Unzip()`can unpack all zip files that are currently selected.

### Fast text editor

The last example of advanced applications in the routines section will be a simple text editor. It is not meant to open text files or save new ones, but simply to edit text on the fly.
```python
import tkinter
import win32.lib.win32con as win32con
import win32.test.test_pywintypes as pywintypes
import win32.win32clipboard as win32clipboard
import win32.win32gui as win32gui

def GuiEditor():
    editor = tkinter.Tk()  # Create GUI
    editor.title("Text Editor")
    editor.rowconfigure(0, minsize=300, weight=1)
    editor.columnconfigure(1, minsize=500, weight=1)

    editorText = tkinter.Text(editor)  # Create text field
    editorText.grid(row=0, column=1, sticky="nsew")

    editor.protocol("WM_DELETE_WINDOW", Gui_Editor_Exit)  # Conditions
    editor.bind("<Escape>", Gui_Editor_Exit)
    editorText.focus_set()  # Fokus window
    editor.after(2, Gui_Editor_Top)  # Keep always in foreground

    editor.mainloop()

def Gui_Editor_Exit(*args):
    win32clipboard.OpenClipboard()  # Save content to clipboard
    win32clipboard.SetClipboardText(editorText.get(1.0, tkinter.END), win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()
    editor.destroy()

def Gui_Editor_Top(): # Keep always in foreground
    win32gui.SetWindowPos(int(editor.frame(), 16), win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
```
With the vanilla package tkinter, a simple text input field can be built that is opened via hotkey and always remains in the foreground. With *Esc* it is immediately closed and the content is copied to the clipboard.

## Conclusion
I have tried to put all my experience from years of working with automation scripts, first in AutoHotkey, then in Python, into this compact guide. Hopefully it was able to help you implement your own automation ideas in Python.

If you have any questions, suggestions for improvement or if you yourself find useful ideas for everyone to present in this guide, please feel free to contact me at info.github@cmuessig.de.

© Christoph Müßig