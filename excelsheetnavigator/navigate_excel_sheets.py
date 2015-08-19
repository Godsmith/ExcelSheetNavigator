from Tkinter import *
import win32com.client
import win32gui # requires pywin32
import unicodedata


class Application(Frame):

    def __init__(self, master):
        Frame.__init__(self, master)
        
        self.worksheets_label_string = StringVar()
        self.pack()
        self.createWidgets()
        
        self.model = NavigatorModel(WindowMgr())
        
        self.update()

    def createWidgets(self):
        self.worksheets_label = Label(self, textvariable = self.worksheets_label_string)
        self.worksheets_label.pack()
    
        self.text_box = Entry(self, width=30)
        self.text_box.pack({"side": "left"})
        self.text_box.focus_set()
        
        self.ok = Button(self)
        self.ok["text"] = "OK"
        self.ok["command"] = self.enter_press
        self.ok.pack({"side": "left"})
        
        self.cancel = Button(self)
        self.cancel["text"] = "Cancel"
        self.cancel["command"] = self.quit
        self.cancel.pack({"side": "left"})
    
    def key_press(self, event):
        if event.char == '\x1b': # Escape
            self.quit()
        elif event.char == '\r': # Enter
            self.enter_press()
        else:
            self.update()
    
    def update(self):
        text_to_show = self.get_text_to_show()
        self.display_string(text_to_show)
        
    def get_text_to_show(self):
        entered_text = self.text_box.get()
        return self.model.text_to_show(entered_text)
    
    def display_string(self, s):
        self.worksheets_label_string.set(s)

    def enter_press(self):
        if self.get_text_to_show() == '':
            return
        self.model.switch_to_first_worksheet_in_list()
        self.quit()
       
    def do_quit(self, event):
        self.quit()

class NavigatorModel:

    def __init__(self, window_mgr):
        self.window_mgr = window_mgr
        excel = win32com.client.Dispatch("Excel.Application")
        self._workbook_from_worksheet_name = {}
        self._worksheet_names = []
        for i in range(1, excel.Workbooks.Count+1):
            workbook = excel.Workbooks(i)
            for j in range(1, workbook.Worksheets.Count+1):
                worksheet = workbook.Worksheets(j)
                self._worksheet_names.append(worksheet.Name)
                self._workbook_from_worksheet_name[worksheet.Name] = workbook
        self._filtered_worksheet_list = []
    
    def text_to_show(self, text_box_text):
        current_words = text_box_text.split(' ')
        self._filtered_worksheet_list = [worksheet_name for worksheet_name in self._worksheet_names if all(word.lower() in worksheet_name.lower() for word in current_words)]
        self._filtered_worksheet_list.sort()
        return '\n'.join(self._filtered_worksheet_list)
        
    def switch_to_first_worksheet_in_list(self):
        worksheet_name = self._filtered_worksheet_list[0]
        workbook = self._workbook_from_worksheet_name[worksheet_name]
        workbook.Worksheets(worksheet_name).Activate()
        self.window_mgr.find_window_text(workbook.Name)
        self.window_mgr.set_foreground()
    

class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""
    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name = None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, text):
        '''Pass to win32gui.EnumWindows() to check all the opened windows'''
        window_text = str(win32gui.GetWindowText(hwnd))
        text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore')
        if text in window_text and "Microsoft Visual Basic for Applications" not in window_text:
            self._handle = hwnd

    def find_window_text(self, text):
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, text)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)


def main():
    
    #print worksheet_names
    root = Tk()
    app = Application(root)
    root.bind('<Key>', app.key_press)
    app.mainloop()
    #root.destroy()

    
if __name__ == "__main__":
    main()
