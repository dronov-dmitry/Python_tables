

def create_dialog():
    localContext = uno.getComponentContext()
    smgr = localContext.getServiceManager()
    
    # Создаем диалог
    dialog = smgr.createInstanceWithContext("com.sun.star.awt.Dialog", localContext)
    dialog.Title = "Мой первый диалог на Python"
    dialog.Modal = True

    # Создаем текстовое поле
    ctrl = smgr.createInstanceWithContext("com.sun.star.awt.TextField", localContext)
    ctrl.Label = "Введите текст:"
    dialog.addControl(ctrl)

    # Создаем кнопку "ОК"
    ctrl = smgr.createInstanceWithContext("com.sun.star.awt.PushButton", localContext)
    ctrl.Label = "ОК"
    dialog.addControl(ctrl)

    # Покажем диалог
    dialog.execute()




import tkinter as tk
from tkinter import simpledialog, messagebox


import tkinter as tk
from tkinter import simpledialog, messagebox


class CustomDialog(simpledialog.Dialog):
    def __init__(self, parent, title=None):
        super().__init__(parent, title)

    def body(self, master):
        """Create dialog body."""
        tk.Label(master, text="Name:").grid(row=0, column=0, sticky="w")
        tk.Label(master, text="Age:").grid(row=1, column=0, sticky="w")

        self.name_var = tk.StringVar()  # Ensure StringVar is properly initialized
        self.age_var = tk.StringVar()

        self.name_entry = tk.Entry(master, textvariable=self.name_var)
        self.age_entry = tk.Entry(master, textvariable=self.age_var)

        self.name_entry.grid(row=0, column=1, padx=100, pady=5)
        self.age_entry.grid(row=1, column=1, padx=100, pady=5)

        return self.name_entry  # Focus on the name entry by default

    def validate(self):
        """Validate the user input."""
        name = self.name_entry.get().strip()  # Ensure we fetch the trimmed value
        age_str = self.age_entry.get().strip()

        try:
            if not name:
                raise ValueError("Name cannot be empty.")
            if not age_str.isdigit():
                raise ValueError("Age must be a whole number.")
            age = int(age_str)
            if age < 0:
                raise ValueError("Age must be a positive number.")
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e))
            return False

        print(f"DEBUG: name_var={name}, age_var={age_str}")  # Debugging print statement
        return True

    def apply(self):
        """Handle user input after validation."""
        self.result = {
            "name": self.name_entry.get().strip(),  # Fetch the trimmed value
            "age": int(self.age_entry.get().strip())
        }


# Main application
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    dialog = CustomDialog(root, title="Enter Your Details")
    # dialog.show()
    

    if dialog.result:
        name = dialog.result["name"]
        age = dialog.result["age"]
        messagebox.showinfo("Info", f"Name: {name}\nAge: {age}")
    else:
        name = ''
        age = ''
        messagebox.showinfo("Info", "Dialog canceled.")
    # root.mainloop()
    return name, age
    
# main()

def hello_world(*args):
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.Sheets[0]
    cell1 = sheet.getCellByPosition(0, 0)  # A1
    cell2 = sheet.getCellRangeByName("A2")
    # create_dialog()
    # main()
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    dialog = CustomDialog(root, title="Enter Your Details")

    if dialog.result:
        name = str(dialog.result["name"])
        age = str(dialog.result["age"])
        messagebox.showinfo("Info", f"Name: {name}\nAge: {age}")
    else:
        name = '1'
        age = '2'
        messagebox.showinfo("Info", "Dialog canceled.")
    # dialog = InputDialog()
    # result = dialog.print_info()
    # messagebox.showinfo("Успешно", str(result))
    # dialog = InputDialog()
    # ctx = uno.getComponentContext()
    # dialog = DialogExample(ctx)
    # dialog.show_dialog()
    ptV = str(sys.version)
    ptVer = str(sys.version_info[1])
    ptExe = str(sys.executable)
    ptExe.replace('/', '-')
    # ptExe = str(type(ptExe[0]))
    # cell.Value = "sdsd"
    # cell_a1 = sheet.getCellRange("A1")
    # cell_a1.String = "sdsd"
    # messagebox.showinfo("Info", cell_a1)
    cell1.setString(name)
    cell2.setValue(int(age))
    return "Hello, World!"  
    # return str(sys.version_info[0])
