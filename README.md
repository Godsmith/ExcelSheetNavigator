# ExcelSheetNavigator
Python script to quickly navigate between sheets in large Excel workbooks

## Dependencies
- Windows
- Excel
- Python 2.7
- [PyWin32](http://sourceforge.net/projects/pywin32/)

## Usage
1. Make sure Excel is open.
2. Start the script.
  1. In order for this to actually save any time, the script should be started by using a [hotkey tied to a Windows shortcut](http://windows.microsoft.com/en-us/windows/create-keyboard-shortcuts-open-programs#1TC=windows-7) or using for example [AutoHotkey](http://www.autohotkey.com/)).
  2. Use pythonw.exe instead of python.exe to avoid the cmd window to pop up.
3. Type a partial case-insensitive name of the sheet to navigate to. Multiple search terms are separated by spaces. For example, "impo tabl" will match "MyImportantTable".
4. Press Enter to switch to the first sheet in the list.
