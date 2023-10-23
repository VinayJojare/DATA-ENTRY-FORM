import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os


window = tkinter.Tk()

window.title("DATA ENTRY FORM")

frame = tkinter.Frame(window)
frame.pack()

def get_data():
    accepted = terms_accept.get()
    if accepted == "accepted":
        
        username =  first_name_entry.get()
        surname = last_name_entry.get()
        title = title_combobox.get()
        age = age_label_spin.get()
        country = country_label_combobox.get()
        Course = course_check.get()
        duration = course_duration_check.get()
        year = course_year_check.get()

        if username and surname and age and country and Course and duration and year:
        
            print("TITLE:", title)
            print("FIRST NAME:", username)
            print("LAST NAME:", surname)
            print("AGE:", age)
            print("COUNTRY:", country)
            print("COURSE:", Course)
            print("DURATION:", duration)
            print("YEAR:", year)
            print("--------------------------------------------------------------------------")
            #_______________________________________________________________________
            # accessing data in Excel sheet
            filepath = r"E:\Python Projects\Data Entry form\entrydata.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                Sheet = workbook.active
                heading = ["TITLE", "FIRST NAME", "LAST NAME", "AGE", "COUNTRY", "COURSE", "DURATION", "YEAR"]
                Sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            Sheet = workbook.active
            data = [title, username, surname, age, country, Course, duration, year]
            Sheet.append(data)
            workbook.save(filepath)
            # _____________________________________________________________

        else:
            tkinter.messagebox.showwarning(title="ERROR", message="please fill all information")
                
        
    else:
        tkinter.messagebox.showwarning(title="ERROR", message="not accepted terms and conditions")



user_info_label = tkinter.LabelFrame(frame, text="USER INFORMATION")
user_info_label.grid(row=0, column=0, sticky="news", padx=10, pady=10)


first_name = tkinter.Label(user_info_label, text="FIRST NAME")
first_name.grid(row=0, column=0)
last_name = tkinter.Label(user_info_label, text="LAST NAME")
last_name.grid(row=0, column=1)


first_name_entry = tkinter.Entry(user_info_label)
first_name_entry.grid(row=1, column=0, padx=5, pady=5)
last_name_entry = tkinter.Entry(user_info_label)
last_name_entry.grid(row=1, column=1, padx=5, pady=5)


title_label = tkinter.Label(user_info_label, text="title")
title_combobox = ttk.Combobox(user_info_label, values=["", "MR.", "MS.", "DR."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)


age_label = tkinter.Label(user_info_label, text="AGE")
age_label_spin = ttk.Spinbox(user_info_label, from_=18, to=100)
age_label.grid(row=2, column=0)
age_label_spin.grid(row=3, column=0)


country_label = tkinter.Label(user_info_label, text="COUNTRY")
country_label_combobox = ttk.Combobox(user_info_label, values=["US","UK","IND","SL","RUS"])
country_label.grid(row=2, column=1, padx=5, pady=5)
country_label_combobox.grid(row=3, column=1, padx=5, pady=5)


course_label = tkinter.LabelFrame(frame, text="COURSE")
course_label.grid(row=1, column=0, sticky="news", padx=10, pady=10)


registration_label = tkinter.Label(course_label, text="Course Completed")
course_check = ttk.Combobox(course_label, values=["Python", "JAVA", "C++", "Javascript"])
registration_label.grid(row=0, column=0, padx=5, pady=5)
course_check.grid(row=1, column=0, padx=5, pady=5)


course_duration = tkinter.Label(course_label, text="Duration")
course_duration_check = ttk.Combobox(course_label, values=["0-6 month", "1 year", "2 year", "3 year"])
course_duration.grid(row=0, column=1)
course_duration_check.grid(row=1, column=1, padx=5, pady=5)

course_year = tkinter.Label(course_label, text="YEAR")
course_year_check = ttk.Combobox(course_label, values=["2017", "2018", "2019", "2020", "2021", "2020", "2021", "2022", "2023"])
course_year.grid(row=0, column=2)
course_year_check.grid(row=1, column=2, padx=5, pady=5)


terms_label = tkinter.LabelFrame(frame, text="TERMS & CONDISION")
terms_label.grid(row=2, column=0, sticky="news", padx=10, pady=10)

terms_accept = tkinter.StringVar(value="not accepted")
terms_label_check = ttk.Checkbutton(terms_label, text="Accept the terms & conditions",variable=terms_accept, onvalue="accepted", offvalue="not accepted")
terms_label_check.grid(row=0, column=0, )


button = tkinter.Button(frame, text="SUBMIT", border=5, command=get_data)
button.grid(row=3, column=0)


window.mainloop()

