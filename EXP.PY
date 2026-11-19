import tkinter as tk
import sqlite3 as sql
import pandas as pd
from tkinter import messagebox, ttk
from datetime import datetime

conn = sql.connect("Expenses.db")
cursor = conn.cursor()
cursor.execute('''CREATE TABLE IF NOT EXISTS expenses (
    ID INTEGER PRIMARY KEY AUTOINCREMENT,
    MONTH TEXT NOT NULL,
    EXPENSES REAL NOT NULL,
    CATEGORY TEXT NOT NULL,
    REMARKS TEXT,
    DATE TEXT NOT NULL
)''')
conn.commit()


def load_data(event=None):
    for row in tree.get_children():
        tree.delete(row)

    selected_month = filter_month.get()
    if selected_month == "All":
        cursor.execute("SELECT * FROM expenses ORDER BY ID DESC")
    else:
        cursor.execute("SELECT * FROM expenses WHERE MONTH=? ORDER BY ID DESC", (selected_month,))

    total = 0
    for r in cursor.fetchall():
        tree.insert("", "end", values=r)
        total += r[2]

    total_lbl.config(text=f"Total Expenses (in Rs.): {total:.2f}")


def add_expense(event=None):
    try:
        expense_value = float(expense.get())
        if expense_value <= 0:
            messagebox.showinfo("Alert", "Expense must be greater than 0!")
            return
    except ValueError:
        messagebox.showinfo("Alert", "Expense must be a number!")
        return

    date_value = datetime.now().strftime("%Y-%m-%d %H:%M")
    cursor.execute(
        "INSERT INTO expenses (MONTH, EXPENSES, CATEGORY, REMARKS, DATE) VALUES (?, ?, ?, ?, ?)",
        (month.get(), expense_value, category.get(), remarks.get().strip(), date_value)
    )
    conn.commit()
    clear_form()
    load_data()
    messagebox.showinfo("Success", "Record inserted successfully!")


def clear_form():
    expense.delete(0, "end")
    remarks.delete(0, "end")
    month.current(0)
    category.current(0)
    expense.focus()


def delete_expense():
    selected = tree.selection()
    if not selected:
        messagebox.showinfo("Alert", "Select a record to delete!")
        return

    if not messagebox.askyesno("Confirm", "Are you sure you want to delete this record?"):
        return

    record_id = tree.item(selected[0])["values"][0]
    cursor.execute("DELETE FROM expenses WHERE ID=?", (record_id,))
    conn.commit()
    load_data()


def export_data():
    cursor.execute("SELECT * FROM expenses")
    data = cursor.fetchall()
    if not data:
        messagebox.showinfo("Alert", "No data to export!")
        return

    df = pd.DataFrame(data, columns=["ID", "MONTH", "EXPENSES", "CATEGORY", "REMARKS", "DATE"])
    df.to_excel("Expenses_Report.xlsx", index=False)
    messagebox.showinfo("Success", "Exported to Expenses_Report.xlsx")


def on_close():
    conn.close()
    root.destroy()


root = tk.Tk()
root.title("Expenses Tracker")
root.geometry("850x620")
root.protocol("WM_DELETE_WINDOW", on_close)

title = tk.Label(root, text="Personal Expenses Tracker", font=("Verdana", 18, "bold"))
title.pack(pady=(15, 5))


frame1 = tk.Frame(root)
frame1.pack(pady=10)

all_months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
all_categories = ["Food", "Makeup Kids", "Travel", "Clothes", "Other"]

tk.Label(frame1, text="Select Month", font=("Verdana", 12)).grid(row=0, column=0, pady=5, padx=5, sticky="w")
month = ttk.Combobox(frame1, values=all_months, state="readonly", font=("Verdana", 12), width=18)
month.grid(row=0, column=1, pady=5)
month.current(0)

tk.Label(frame1, text="Enter Expense (Rs)", font=("Verdana", 12)).grid(row=1, column=0, pady=5, padx=5, sticky="w")
expense = tk.Entry(frame1, width=20, font=("Verdana", 12))
expense.grid(row=1, column=1, pady=5)

tk.Label(frame1, text="Select Category", font=("Verdana", 12)).grid(row=2, column=0, pady=5, padx=5, sticky="w")
category = ttk.Combobox(frame1, values=all_categories, state="readonly", font=("Verdana", 12), width=18)
category.grid(row=2, column=1, pady=5)
category.current(0)

tk.Label(frame1, text="Enter Spent Item", font=("Verdana", 12)).grid(row=3, column=0, pady=5, padx=5, sticky="w")
remarks = tk.Entry(frame1, width=20, font=("Verdana", 12))
remarks.grid(row=3, column=1, pady=5)

root.bind("<Return>", add_expense)


btn_frame = tk.Frame(root)
btn_frame.pack(pady=5)

btn_style = {"width": 10, "font": ("Verdana", 11, "bold"), "cursor": "hand2"}
tk.Button(btn_frame, text="Save", bg="green", fg="white", command=add_expense, **btn_style).grid(row=0, column=0, padx=6)
tk.Button(btn_frame, text="Clear", bg="orange", fg="black", command=clear_form, **btn_style).grid(row=0, column=1, padx=6)
tk.Button(btn_frame, text="Delete", bg="red", fg="white", command=delete_expense, **btn_style).grid(row=0, column=2, padx=6)
tk.Button(btn_frame, text="Export", bg="blue", fg="white", command=export_data, **btn_style).grid(row=0, column=3, padx=6)


info_frame = tk.Frame(root)
info_frame.pack(pady=10)

tk.Label(info_frame, text="Filter by Month:", font=("Verdana", 12)).grid(row=0, column=0, padx=5)
filter_month = ttk.Combobox(info_frame, values=["All"] + all_months, state="readonly", font=("Verdana", 12), width=8)
filter_month.grid(row=0, column=1, padx=5)
filter_month.current(0)
filter_month.bind("<<ComboboxSelected>>", load_data)

total_lbl = tk.Label(info_frame, text="Total Expenses (in Rs.): 0", font=("Verdana", 13, "bold"))
total_lbl.grid(row=0, column=2, padx=30)


frame2 = tk.Frame(root)
frame2.pack(pady=10, padx=15, fill="both", expand=True)

tree = ttk.Treeview(frame2, columns=("ID", "Month", "Expense", "Category", "Remarks", "Date"), show="headings")
for col, width in [("ID", 40), ("Month", 80), ("Expense", 90), ("Category", 90), ("Remarks", 170), ("Date", 130)]:
    tree.heading(col, text=col.upper())
    tree.column(col, width=width, anchor="center")

scrollbar = ttk.Scrollbar(frame2, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
tree.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

load_data()
expense.focus()
root.mainloop()

