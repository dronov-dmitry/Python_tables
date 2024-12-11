import xlwings as xw
import tkinter as tk
from tkinter import simpledialog, messagebox

# Custom dialog class
class CustomDialog(simpledialog.Dialog):
    def __init__(self, parent, title=None):
        super().__init__(parent, title)
    
    def body(self, master):
        """Create dialog body."""
        tk.Label(master, text="Name:").grid(row=0, column=0, sticky="w")
        tk.Label(master, text="Age:").grid(row=1, column=0, sticky="w")
        
        self.name_var = tk.StringVar()
        self.age_var = tk.StringVar()
        
        self.name_entry = tk.Entry(master, textvariable=self.name_var)
        self.age_entry = tk.Entry(master, textvariable=self.age_var)
        
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)
        self.age_entry.grid(row=1, column=1, padx=5, pady=5)
        
        return self.name_entry  # Focus on the name entry by default
    
    def validate(self):
        """Validate the user input."""
        try:
            age = int(self.age_var.get())
            if not self.name_var.get().strip():
                raise ValueError("Name cannot be empty.")
            if age < 0:
                raise ValueError("Age must be a positive number.")
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e))
            return False
        return True
    
    def apply(self):
        """Handle user input after validation."""
        self.result = {
            "name": self.name_var.get(),
            "age": int(self.age_var.get())
        }

# Main application
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    dialog = CustomDialog(root, title="Enter Your Details")
    if dialog.result:
        name = dialog.result["name"]
        age = dialog.result["age"]
        messagebox.showinfo("Info", f"Name: {name}\nAge: {age}")
    else:
        messagebox.showinfo("Info", "Dialog canceled.")

def my_macro():
    # Открываем активную книгу
    wb = xw.Book.caller()
    main()
    sheet = wb.sheets[0]  # Используем первый лист

    # Изменяем ячейку A1
    sheet["A1"].value = "Привет, Excel!"