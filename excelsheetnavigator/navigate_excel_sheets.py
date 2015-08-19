from Tkinter import *
import win32com.client
import win32gui # requires pywin32
import unicodedata


class Application(Frame):

    def __init__(self, master, workbook_from_worksheet_names):
        Frame.__init__(self, master)

        self.workbook_from_worksheet_names = workbook_from_worksheet_names
        self.worksheet_names = workbook_from_worksheet_names.keys()
        
        self.worksheets_label_string = StringVar()



        self.pack()
        self.createWidgets()

        self.show_filtered_worksheet_list()

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
            self.show_filtered_worksheet_list()

    def show_filtered_worksheet_list(self):
        current_words = self.text_box.get().split(' ')

        self.filtered_worksheet_list = [worksheet_name for worksheet_name in self.worksheet_names if all(word.lower() in worksheet_name.lower() for word in current_words)]
        self.filtered_worksheet_list.sort()
        self.worksheets_label_string.set('\n'.join(self.filtered_worksheet_list))

    
    def enter_press(self):
        if len(self.filtered_worksheet_list) == 0:
            return
        worksheet_name = self.filtered_worksheet_list[0]
        workbook = self.workbook_from_worksheet_names[worksheet_name]
        workbook.Worksheets(worksheet_name).Activate()
        w = WindowMgr()
        w.find_window_text(workbook.Name)
        w.set_foreground()

        self.quit()
       
    def do_quit(self, event):
        self.quit()

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
        #if u'cand_epg (1).xls' in window_text:
        #print text
         #   print window_text
        text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore')
        if text in window_text and "Microsoft Visual Basic for Applications" not in window_text:
            self._handle = hwnd

    def find_window_text(self, text):
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, text)

    def set_foreground(self):
        """put the window in the foreground"""
        print self._handle
        win32gui.SetForegroundWindow(self._handle)


def main():


    excel = win32com.client.Dispatch("Excel.Application")
    workbook_from_worksheet_name = {}
    worksheet_names = []
    for i in range(1, excel.Workbooks.Count+1):
        workbook = excel.Workbooks(i)
        for j in range(1, workbook.Worksheets.Count+1):
            worksheet = workbook.Worksheets(j)
            worksheet_names.append(worksheet.Name)
            workbook_from_worksheet_name[worksheet.Name] = workbook

    #print worksheet_names
    root = Tk()
    app = Application(root, workbook_from_worksheet_name)
    root.bind('<Key>', app.key_press)
    app.mainloop()
    #root.destroy()



if __name__ == "__main__":
    main()
