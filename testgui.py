import tkinter as tk
from tkinter import ttk
import pandas as pd

# Load CSV data
df = pd.read_csv('data.csv')

# Create GUI window
root = tk.Tk()
root.title("CSV Table")

# Create Treeview widget
tree = ttk.Treeview(root)
tree["columns"] = list(df.columns)
tree["show"] = "headings"  # Hide the default first column

# Add headings
for col in df.columns:
    tree.heading(col, text=col)

# Insert data rows
for index, row in df.iterrows():
    tree.insert("", "end", values=list(row))

tree.pack(expand=True, fill='both')

# Run the GUI
root.mainloop()
