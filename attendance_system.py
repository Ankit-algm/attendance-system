"""
Smart Attendance System
Class XII Informatics Practices Project - Army Public School, ASC C & C, Bangalore

A Tkinter desktop app that lets a teacher mark daily attendance, manage
student records and marks, view charts, and export a PDF report.
Data is stored in Excel files using pandas.
"""

import pandas as pd
from datetime import date
import os
from tkinter import *
from tkinter import messagebox, simpledialog, ttk, filedialog
import matplotlib.pyplot as plt
from fpdf import FPDF

# ------------------- CONFIG -------------------
STUDENT_FILE = "students.xlsx"
ATTENDANCE_PREFIX = "attendance_"
PASSWORD = "teacher123"


# ------------------- INITIAL SETUP -------------------
def init_student_file():
    if not os.path.exists(STUDENT_FILE):
        df = pd.DataFrame(columns=["Roll No", "Name", "Marks"])
        initial = [
            (1, "Nikhil", 99), (2, "Ankit", 99), (3, "Shreya", 99),
            (4, "Aarav", 92), (5, "Aditi", 70), (6, "Aryan", 88),
            (7, "Ananya", 65), (8, "Dev", 95), (9, "Diya", 80),
            (10, "Ishaan", 76), (11, "Meera", 82), (12, "Riya", 69)
        ]
        for r, n, m in initial:
            df.loc[len(df)] = [r, n, m]
        df.to_excel(STUDENT_FILE, index=False)


def load_students():
    return pd.read_excel(STUDENT_FILE)


def save_students(df):
    df.to_excel(STUDENT_FILE, index=False)


def get_month_file(month_name=None):
    if month_name is None:
        month_name = date.today().strftime("%B_%Y")
    file = f"{ATTENDANCE_PREFIX}{month_name}.xlsx"
    if not os.path.exists(file):
        pd.DataFrame(columns=["Date", "Roll No", "Name", "Status"]).to_excel(file, index=False)
    return file


def load_data(file):
    return pd.read_excel(file)


def save_data(df, file):
    df.to_excel(file, index=False)


# ------------------- ATTENDANCE FUNCTIONS -------------------
def mark_attendance():
    file = get_month_file()
    df = load_data(file)
    today = date.today().strftime("%d-%m-%Y")
    if today in df["Date"].values:
        messagebox.showwarning("Already Marked", f"Attendance already marked for {today}.")
        return

    df_students = load_students()
    win = Toplevel(root)
    win.title("Mark Attendance")
    frame = Frame(win)
    frame.pack(fill=BOTH, expand=True)
    canvas = Canvas(frame)
    scroll = Scrollbar(frame, orient=VERTICAL, command=canvas.yview)
    inner = Frame(canvas)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scroll.set)
    canvas.pack(side=LEFT, fill=BOTH, expand=True)
    scroll.pack(side=RIGHT, fill=Y)

    status_vars = {}
    for _, st in df_students.iterrows():
        var = BooleanVar(value=True)
        chk = Checkbutton(inner, text=f"{st['Roll No']} - {st['Name']}", variable=var)
        chk.pack(anchor="w", padx=5, pady=2)
        status_vars[st["Roll No"]] = (st["Name"], var)

    def save():
        for roll, (name, var) in status_vars.items():
            df.loc[len(df)] = [today, roll, name, "Present" if var.get() else "Absent"]
        save_data(df, file)
        messagebox.showinfo("Saved", f"Attendance saved for {today}")
        win.destroy()

    Button(win, text="Save Attendance", command=save).pack(pady=8)


def view_monthly():
    files = [f for f in os.listdir() if f.startswith(ATTENDANCE_PREFIX)]
    if not files:
        messagebox.showinfo("No Files", "No attendance files found.")
        return
    file = filedialog.askopenfilename(title="Select attendance file",
                                      filetypes=[("Excel files", "*.xlsx")])
    if not file:
        return
    df = pd.read_excel(file)
    show_table(df, f"Attendance for {os.path.basename(file)}")


def attendance_summary():
    files = [f for f in os.listdir() if f.startswith(ATTENDANCE_PREFIX)]
    if not files:
        messagebox.showerror("Missing", "No attendance files.")
        return
    all_data = pd.concat([pd.read_excel(f) for f in files], ignore_index=True)
    summary = all_data.groupby(["Roll No", "Name"])["Status"].apply(
        lambda s: (s == "Present").sum() / len(s) * 100).reset_index(name="Attendance %")
    show_table(summary, "Attendance Summary")

    plt.figure(figsize=(10, 5))
    colors = ['red' if p < 75 else 'green' for p in summary["Attendance %"]]
    plt.bar(summary["Name"], summary["Attendance %"], color=colors)
    plt.axhline(y=75, color='orange', linestyle='--')
    plt.title("Attendance % of Students")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()


def show_table(df, title):
    win = Toplevel(root)
    win.title(title)
    tree = ttk.Treeview(win, columns=list(df.columns), show="headings")
    for c in df.columns:
        tree.heading(c, text=c)
        tree.column(c, width=120, anchor="center")
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))
    scrollbar = Scrollbar(win, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=RIGHT, fill=Y)
    tree.pack(side=LEFT, fill=BOTH, expand=True)


# ------------------- STUDENT MANAGEMENT -------------------
def add_student():
    df = load_students()
    roll = simpledialog.askinteger("Add", "Roll No:")
    if roll is None:
        return
    if roll in df["Roll No"].values:
        messagebox.showerror("Error", "Roll number exists.")
        return
    name = simpledialog.askstring("Add", "Name:")
    marks = simpledialog.askinteger("Add", "Marks:")
    if not name or marks is None:
        return
    df.loc[len(df)] = [roll, name, marks]
    save_students(df)
    messagebox.showinfo("Added", f"{name} added.")


def remove_student():
    df = load_students()
    roll = simpledialog.askinteger("Remove", "Enter Roll No:")
    if roll is None:
        return
    if roll not in df["Roll No"].values:
        messagebox.showerror("Error", "Not found.")
        return
    df = df[df["Roll No"] != roll]
    save_students(df)
    messagebox.showinfo("Removed", f"Removed roll {roll}")


def modify_marks():
    df = load_students()
    roll = simpledialog.askinteger("Modify", "Enter Roll No:")
    if roll is None:
        return
    if roll not in df["Roll No"].values:
        messagebox.showerror("Error", "Not found.")
        return
    new_marks = simpledialog.askinteger("Modify", "Enter new marks:")
    if new_marks is None:
        return
    df.loc[df["Roll No"] == roll, "Marks"] = new_marks
    save_students(df)
    messagebox.showinfo("Updated", f"Marks updated for {roll}")


def marks_analysis():
    df = load_students()
    show_table(df, "Marks Analysis")
    plt.figure(figsize=(10, 5))
    plt.bar(df["Name"], df["Marks"], color="skyblue")
    plt.title("Marks Performance")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


# ------------------- DELETE & EXPORT -------------------
def delete_all():
    if messagebox.askyesno("Confirm", "Delete all attendance data?"):
        for f in os.listdir():
            if f.startswith(ATTENDANCE_PREFIX):
                os.remove(f)
        messagebox.showinfo("Deleted", "All attendance files removed.")


def export_report():
    df = load_students()
    files = [f for f in os.listdir() if f.startswith(ATTENDANCE_PREFIX)]
    if files:
        all_data = pd.concat([pd.read_excel(f) for f in files], ignore_index=True)
        summary = all_data.groupby(["Roll No", "Name"])["Status"].apply(
            lambda s: (s == "Present").sum() / len(s) * 100).reset_index(name="Attendance %")
    else:
        summary = pd.DataFrame(columns=["Roll No", "Name", "Attendance %"])

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=14)
    pdf.cell(200, 10, txt="Attendance & Marks Report", ln=True, align="C")
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="\nAttendance Summary\n", ln=True)
    for _, row in summary.iterrows():
        pdf.cell(200, 8, txt=f"{row['Roll No']} - {row['Name']}: {row['Attendance %']:.2f}%", ln=True)
    pdf.cell(200, 10, txt="\nMarks Summary\n", ln=True)
    for _, row in df.iterrows():
        pdf.cell(200, 8, txt=f"{row['Roll No']} - {row['Name']}: {row['Marks']} marks", ln=True)

    save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF files", "*.pdf")])
    if save_path:
        pdf.output(save_path)
        messagebox.showinfo("Exported", f"Report saved at {save_path}")


# ------------------- LOGIN SYSTEM -------------------
def login():
    pwd = simpledialog.askstring("Login", "Enter teacher password:", show="*")
    if pwd != PASSWORD:
        messagebox.showerror("Denied", "Wrong password!")
        root.destroy()
        raise SystemExit  # stop here so the rest of the program does not run
    else:
        messagebox.showinfo("Welcome", "Access granted")


# ------------------- GUI LAYOUT -------------------
root = Tk()
root.title("Attendance & Marks System")
root.geometry("450x700")
login()
init_student_file()

Label(root, text="Attendance & Marks System", font=("Arial", 16, "bold")).pack(pady=10)
Button(root, text="Mark Attendance", width=30, command=mark_attendance).pack(pady=5)
Button(root, text="View Monthly Attendance", width=30, command=view_monthly).pack(pady=5)
Button(root, text="Attendance Summary", width=30, command=attendance_summary).pack(pady=5)
Button(root, text="Marks Analysis", width=30, command=marks_analysis).pack(pady=5)

Label(root, text="\nStudent Management", font=("Arial", 13, "bold")).pack(pady=6)
Button(root, text="Add Student", width=30, command=add_student).pack(pady=5)
Button(root, text="Remove Student", width=30, command=remove_student).pack(pady=5)
Button(root, text="Modify Marks", width=30, command=modify_marks).pack(pady=5)

Label(root, text="\nFile Management", font=("Arial", 13, "bold")).pack(pady=6)
Button(root, text="Export Report (PDF)", width=30, command=export_report).pack(pady=5)
Button(root, text="Delete All Attendance Data", width=30, command=delete_all).pack(pady=5)

Label(root, text="").pack()
Button(root, text="Exit", width=30, command=root.destroy).pack(pady=10)

root.mainloop()
