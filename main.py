import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap import ttk
import pyodbc
import pandas as pd

class AccessToCSVApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Access Database to CSV")
        self.root.geometry("600x400")

        self.style = Style(theme="flatly")  # Apply a modern theme
        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        label_file_path = ttk.Label(frame, text="Access Database File:")
        label_file_path.grid(row=0, column=0, sticky=tk.W, pady=5)

        self.entry_file_path = ttk.Entry(frame, width=50)
        self.entry_file_path.grid(row=0, column=1, pady=5, sticky=tk.EW)

        button_browse = ttk.Button(frame, text="Browse...", command=self.select_file, bootstyle="primary-outline")
        button_browse.grid(row=0, column=2, padx=5, pady=5)

        label_query = ttk.Label(frame, text="SQL Query:")
        label_query.grid(row=1, column=0, sticky=tk.NW, pady=5)

        self.text_query = tk.Text(frame, width=50, height=10)
        self.text_query.grid(row=1, column=1, columnspan=2, pady=5, sticky=tk.EW)

        button_generate = ttk.Button(frame, text="Generate CSV", command=self.generate_csv, bootstyle="success-outline")
        button_generate.grid(row=2, column=1, pady=10)

        self.status_label = ttk.Label(frame, text="", foreground="blue")
        self.status_label.grid(row=3, column=0, columnspan=3, pady=5)

        frame.columnconfigure(1, weight=1)

        for widget in frame.winfo_children():
            widget.grid_configure(padx=5, pady=5)

        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat")
        style.configure("TLabel", padding=6)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Access Database Files", "*.accdb *.mdb")])
        if file_path:
            self.entry_file_path.delete(0, tk.END)
            self.entry_file_path.insert(0, file_path)

    def generate_csv(self):
        file_path = self.entry_file_path.get()
        query = self.text_query.get("1.0", tk.END).strip()

        if not file_path or not query:
            messagebox.showerror("Error", "Please provide both file path and SQL query.")
            return

        try:
            self.status_label.config(text="Processing...")
            conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={file_path};'
            conn = pyodbc.connect(conn_str)
            data = pd.read_sql(query, conn)
            conn.close()

            save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                     filetypes=[("CSV Files", "*.csv")])
            if save_path:
                data.to_csv(save_path, index=False)
                messagebox.showinfo("Success", f"CSV file has been saved to {save_path}")
                self.status_label.config(text="CSV file saved successfully.")

        except pyodbc.Error as db_err:
            messagebox.showerror("Database Error", f"An error occurred with the database: {db_err}")
            self.status_label.config(text="")
        except pd.io.sql.DatabaseError as sql_err:
            messagebox.showerror("SQL Error", f"An SQL error occurred: {sql_err}")
            self.status_label.config(text="")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_label.config(text="")

if __name__ == "__main__":
    root = tk.Tk()
    app = AccessToCSVApp(root)
    root.mainloop()
