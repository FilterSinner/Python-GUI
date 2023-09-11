# Employee Name, Dept,Employee Number,Signature,Verified Date
#Items,Serial No,Remarks

import tkinter
from tkinter import Button, ttk
from tkcalendar import Calendar,DateEntry
import openpyxl
import os
from tkinter import filedialog
from PIL import Image, ImageTk


def enter_data():
    accepted = accepted_var.get()
    if accepted == "Accepted":
        #User Info
        name=name_label_e.get()
        empno=EmpNo_label_e.get()
        dept=Dept_label_e.get()
        title=title_label_combobox.get()
        email=Email_label_e.get()
        #Items
        items=Items_label_combobox.get()
        serialno=Serial_no_label_e.get()
        remarks=Remarks_label_e.get()
        image_path=selected_image_path.get()

        print(f"{title} Name: ",name)
        print("Empno: ",empno)
        print("Department: ",dept)
        print("Date: ",get_date())
        print("Email Address: ",email)
        print("------------------------")
        print("Item no:",items)
        print("Serial Number:",serialno)
        print("Remarks: ",remarks)
    




        filepath = r'C:\Users\Ojal\Documents\IT asset mngr\main\data.xlsx'
        if not os.path.exists(filepath):
            workbook = openpyxl.Workbook()
            sheet=workbook.active
            heading=["Employee Name", "Deptartment","Emp No","Verified Date","Email","Items","Serial No","Remarks","Signature"]
            sheet.append(heading)
            workbook.save(filepath)
        workbook=openpyxl.load_workbook(filepath)
        sheet=workbook.active
        selected_date = cal.get()

        sheet.append([name,dept,empno,selected_date,email,items,serialno,remarks,image_path])
        workbook.save(filepath)
    else:
        print("Make sure to check the terms and conditions block :/")




window= tkinter.Tk()

#parent window for everything else, big window
window.title("Data Entry for IT Assets")


frame=tkinter.Frame(window,padx=5,pady=20) #creating a frame inside the window
frame.pack() #layout managers(pack,grid)

#Label frames, sort of like groups that come under a frame(think in figma terms)
user_info_frame=tkinter.LabelFrame(frame,text="User Information")
user_info_frame.grid(row=0,column=0,padx=20,pady=20)

name_label=tkinter.Label(user_info_frame,text="Employee Name")
name_label.grid(row=0,column=0)

EmpNo_label=tkinter.Label(user_info_frame,text="Employee Number")
EmpNo_label.grid(row=0,column=1)

name_label_e= tkinter.Entry(user_info_frame)
EmpNo_label_e=tkinter.Entry(user_info_frame)
name_label_e.grid(row=1,column=0)
EmpNo_label_e.grid(row=1,column=1)

title_label= tkinter.Label(user_info_frame,text="Title")
title_label_combobox= ttk.Combobox(user_info_frame,values=["","Mr","Mrs."])
title_label.grid(row=0,column=2)
title_label_combobox.grid(row=1,column=2)

Dept_label=tkinter.Label(user_info_frame,text="Department")
Dept_label.grid(row=2,column=0)
Dept_label_e=tkinter.Entry(user_info_frame)
Dept_label_e.grid(row=3,column=0)

'''Date_label=tkinter.Label(user_info_frame,text="Verified Date (dd/mm/yyyy)")
    Date_label.grid(row=2,column=1)
    Date_label_e=tkinter.Entry(user_info_frame)
    Date_label_e.grid(row=3,column=1)'''

def get_date():
    selected_date = cal.get()
    print(f"Selected date: {selected_date}")

date_var=tkinter.StringVar()
cal = DateEntry(user_info_frame, date_pattern="dd-mm-yyyy",variable=date_var)
cal.grid(row=3,column=1)
'''btn = Button(user_info_frame, text="Sumbit", command=get_date)
btn.grid(row=4,column=1)'''
Date_label=tkinter.Label(user_info_frame,text="Verified Date")
Date_label.grid(row=2,column=1)

Email_label=tkinter.Label(user_info_frame,text="Email Address")
Email_label.grid(row=2,column=2)
Email_label_e=tkinter.Entry(user_info_frame)
Email_label_e.grid(row=3,column=2)


for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=20,pady=5)


#Item frame
item_info_frame=tkinter.LabelFrame(frame,text="Item")
item_info_frame.grid(row=1,column=0,padx=20,pady=20,sticky="news")

Items_label=tkinter.Label(item_info_frame,text="Items")
Items_label.grid(row=0,column=0)
Items_label_combobox= ttk.Combobox(item_info_frame,values=["Laptop","Desktop","Scanner","Printer","Mobile Phone"])
Items_label_combobox.grid(row=1,column=0)

title_label= tkinter.Label(user_info_frame,text="Title")
title_label_combobox= ttk.Combobox(user_info_frame,values=["","Mr","Mrs."])
title_label.grid(row=0,column=2)
title_label_combobox.grid(row=1,column=2)

Serial_no_label=tkinter.Label(item_info_frame,text="Serial No")
Serial_no_label.grid(row=0,column=1)
Serial_no_label_e=tkinter.Entry(item_info_frame)
Serial_no_label_e.grid(row=1,column=1)

Remarks_label=tkinter.Label(item_info_frame,text="Remarks")
Remarks_label.grid(row=0,column=2)
Remarks_label_e=tkinter.Entry(item_info_frame)
Remarks_label_e.grid(row=1,column=2)

for widget in item_info_frame.winfo_children():
    widget.grid_configure(padx=20,pady=5)

selected_image_path = tkinter.StringVar()

def open_image_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png *.jpg *.jpeg *.gif *.bmp")])
    if file_path:
        selected_image_path.set(file_path)
        load_image(file_path)

def load_image(file_path):
    image = Image.open(file_path)
    photo = ImageTk.PhotoImage(image)
    image_label.config(image=photo)
    image_label.image = photo

image_label = tkinter.Label(item_info_frame)


def load_image(file_path):
    image = Image.open(file_path)
    photo = ImageTk.PhotoImage(image)

    # Update the label with the loaded image
    image_label.config(image=photo)
    image_label.image = photo


upload_image_button = tkinter.Button(item_info_frame, text="Upload Signature", command=open_image_dialog)
upload_image_button.grid(row=1, column=4)

terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, padx=20, pady=10)

name=name_label_e.get()
accepted_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text= f"I,{name} hereby acknowledge that above mentioned assets are under my custody.I understand that this asset belongs to\n The Coca-Cola Bottling Company of Bahrain(B.S.C) and is under my possession for carrying out my office work.\n I hereby assure that  I will take care of the assets of the compay to the best possible extend." ,
                                  variable=accepted_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)


button=tkinter.Button(frame,text="Enter data",command=enter_data)
button.grid(row=3,column=0,sticky="news",padx=20,pady=5)


window.mainloop()
#windox box will appear and run until closed by user