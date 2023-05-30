import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl

def enter_data():
    accepted = accept_var.get()

    if accepted=="Accepted":

     firstname = first_name_entry.get()
     lastname = last_name_entry.get()

     if firstname and lastname:
        title = title_combobox.get()
        age = age_spinbox.get()
        nationality = nationality_combox.get()

    # courses info
        registration_status = reg_status_var.get()
        numcourses = numcourses_spinbox.get()
        semesters = semesters_spinbox.get()

        print("First name: ", firstname, "Last name ", lastname)
        print("Title: ", title, "Age: ", age, "Nationality: ", nationality)
        print("# Courses: ", numcourses, "# Semesters: ", semesters)
        print("Registration Status: ", registration_status)
        print("---------------------------------------------")

        # here we create an Excel sheet
        filepath = "Pasta1.xlsx"
        if not os.path.exists(filepath):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            heading = ["Fist Name", "Last Name", "Title", "Age", "Nationality", "# Courses", "# Semesters", "Rgistration_status"]
            sheet.append(heading)
            workbook.save(filepath)
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        sheet.append([firstname, lastname, title, age, nationality, numcourses, semesters, registration_status])
        workbook.save(filepath)


     else:
         tkinter.messagebox.showwarning(title="Name Erro", message="Put yours First name and last name")
    else:
        tkinter.messagebox.showwarning(title="Error temers", message="You not accepted the terms")
Window = tkinter.Tk()
Window.title("Data Entry Form")

frame = tkinter.Frame(Window)
frame.pack()

user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms.", "Dr."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combox = ttk.Combobox(user_info_frame, values=["Europa", "Afrika", "Asia", "Amerika", "Hell", "oceania", "South-amerika", "north-amerika"])
nationality_label.grid(row=2, column=1)
nationality_combox.grid(row=3, column=1)

# here we are fixing the padding of all child elements
for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# saving as courses info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tkinter.Label(courses_frame, text="Registration Status")

reg_status_var = tkinter.StringVar()
registered_check = tkinter.Checkbutton(courses_frame, text="Currently Registration",
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not registered")


registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = tkinter.Label(courses_frame, text="# Completed Courses")
numcourses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

semesters_label = tkinter.Label(courses_frame, text="# Semesters")
semesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
semesters_label.grid(row=0, column=2)
semesters_spinbox.grid(row=1, column=2)

# repair the padding
for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

terms_frame = tkinter.LabelFrame(frame, text="Terms and Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

button = tkinter.Button(frame, text="Enter Data", command=enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

Window.mainloop()
