import tkinter as tk 
from tkinter import *
from tkinter import messagebox
from tkinter import Tk,Label,PhotoImage,ttk
from tkinter import Listbox, Scrollbar, HORIZONTAL, VERTICAL
from time import *
import sys
import sqlite3
import datetime
#ส่ง รูป
from tkinter import filedialog
from PIL import Image, ImageTk #ติดตั้ง pip install Pillow ใน commandก่อน
import os
#ส่ง รูป
from tkinter import filedialog
#ส่งไฟล์ เป็นpdf
import fitz  # pip install PyMuPDF
global global_photo  
global selected_image_label  
from io import BytesIO
# ส่ง email 
from email.message import EmailMessage
import ssl
import smtplib
global search_clicked
# สร้างpdf
from reportlab.pdfgen import canvas  #pip install reportlab
from reportlab.lib.pagesizes import letter
from fpdf import FPDF   #pip install fpdf

def info_home():
    io=StringVar()
    def account():
        code=StringVar()
        def logcode():
            search_clicked = 0
            user_code = user_log.get()  
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            try:
                find_user = ('SELECT * FROM account WHERE code = ?')
                cursor.execute(find_user, [user_code])
                result = cursor.fetchone()
                if result:
                    code, name, age, department = result
                    if  search_clicked<1:
                        listbox_code.delete(0, tk.END)
                        listbox_name.delete(0, tk.END)
                        listbox_age.delete(0, tk.END)
                        listbox_depart.delete(0, tk.END)
                    listbox_code.insert(tk.END, f"รหัสพนักงาน: {code}")
                    listbox_name.insert(tk.END, f"ชื่อ-สกุล: {name}")
                    listbox_age.insert(tk.END, f"อายุ: {age}")
                    listbox_depart.insert(tk.END, f"แผนก: {department}")
                    search_clicked += 1
                    display_image()
                    showrecord(listbox_record) 
                else:
                    popup.grab_set()
                    messagebox.showerror('Error', 'ไม่พบข้อมูลบัญชี')
                    popup.grab_release()
                user_log.delete(0,END)
            except sqlite3.Error as e:
                popup.grab_set()
                messagebox.showerror('Error', 'เกิดข้อผิดพลาดในการอ่านข้อมูลบัญชี: ' + str(e))
                popup.grab_release()
        def display_image():
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            user_code = user_log.get()
            try:
                cursor.execute("SELECT data FROM img WHERE code=?", (user_code,))
                pic = cursor.fetchone()
                if pic and pic[0]:
                    try:
                        search_clicked = 0
                        if search_clicked <1:
                            for widget in productsshow.winfo_children():
                                widget.destroy()
                        for  x in enumerate(pic):
                            image = Image.open(BytesIO(x[1]))
                            target_width, target_height = 135, 135
                            image = image.resize((target_width, target_height))
                            image = ImageTk.PhotoImage(image)

                            img_label = Label(productsshow, image=image, bg='blue',bd=10)
                            img_label.image = image
                            img_label.pack()
                        search_clicked += 1
                    except Exception as e:
                        messagebox.showerror(f"Error displaying image: {str(e)}")
                else:
                    popup.grab_set()
                    messagebox.showerror('No image found for the given code.',"No image")
                    popup.grab_release()
            except sqlite3.Error as e:
                messagebox.showerror('Error', 'เกิดข้อผิดพลาดในการอ่านข้อมูลบัญชี: ' + str(e))
        def showrecord(listbox_record):
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            user_code = user_log.get()
            search_clicked = 0
            total_income = 0
            id=0
            try:
                cursor.execute('SELECT * FROM record WHERE code = ? ORDER BY time', (user_code,))
                results = cursor.fetchall()
                if results:
                    for result in results:
                        id,inout, code, name, depart, time,late, income = result

                        if inout and code and name and depart and time :
                            if search_clicked < 1:
                                listbox_record.delete(0, tk.END)
                            listbox_record.insert(tk.END, f"{inout}, รหัสพนักงาน: {code}, ชื่อ-สกุล: {name}, แผนก: {depart}, เวลา: {time} {late}")
                            if income:
                                total_income += int(income)
                            search_clicked += 1
                        elif income and time:
                            if search_clicked < 1:
                                listbox_record.delete(0, tk.END)
                            listbox_record.insert(tk.END, f"รายได้ประจำวัน วันที่ {time}= {income}")
                            search_clicked += 1
                            user_log.delete(0, tk.END)
                else:
                    popup.grab_set()
                    messagebox.showerror('Error', 'ไม่พบข้อมูลการเข้า-ออก')
                    popup.grab_release()
            except sqlite3.Error as e:
                messagebox.showerror('Error', 'เกิดข้อผิดพลาดในการอ่านข้อมูลบัญชี: ' + str(e))
        def calculate_total_income():
            user_code = user_log.get()  
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            try:
                cursor.execute('SELECT SUM(income) FROM record WHERE code = ?', (user_code,))
                total_income = cursor.fetchone()[0]
                conn.close()
                if total_income is not None:
                    return total_income
                else:
                    return 0  
            except sqlite3.Error as e:
                conn.close()
                popup.grab_set
                messagebox.showerror('Error', 'เกิดข้อผิดพลาดในการคำนวณรายได้: ' + str(e))
                popup.grab_release()
                return 0
        def calculatedisplay():
            total_income = calculate_total_income()
            listbox_cal.delete(0, tk.END) 
            total_income = round(total_income, 3)
            listbox_cal.insert(tk.END, f"{total_income} บาท")
        def edit():
            age=StringVar()
            depart=StringVar()
            name=StringVar()
            code=StringVar()
            def update_data():
                new_code = userup1_code.get()
                new_name = userup_name.get()
                new_age = user_age.get()
                new_depart = user_depart.get()
                conn = sqlite3.connect(r'mydatabaseinout.db')
                cursor = conn.cursor()
                try:
                    cursor.execute('UPDATE account SET name=?, age=?, depart=? WHERE code=?', (new_name, new_age, new_depart, new_code))
                    cursor.execute('UPDATE record SET name=?, depart=? WHERE code=?', (new_name,new_depart, new_code))
                    conn.commit()
                    conn.close()
                    popup.grab_set()
                    messagebox.showinfo("อัปเดตข้อมูล", "ข้อมูลถูกอัปเดตเรียบร้อยแล้ว")
                    popup.grab_release()

                    userup1_code.destroy()
                    userup_name.destroy()
                    user_age.destroy()
                    user_depart.destroy()
                    btedit.destroy()
                except sqlite3.Error as e:
                    conn.rollback()
                    conn.close()
                    popup.grab_set()
                    messagebox.showerror("ข้อผิดพลาด", "เกิดข้อผิดพลาดในการอัปเดตข้อมูล: " + str(e))
                    popup.grab_release()
            def onc_enter(e):
                userup1_code.delete(0,'end')
            def onc_leave(e):
                entry1=userup1_code.get()
                if entry1=='':
                    userup1_code.insert(0,'กรุณากรอกรหัสพนักงาน')
            userup1_code = tk.Entry(popup, width=20,bd=2 ,fg='black', border=0, bg='White', textvariable=code)
            userup1_code.place(x=395, y=387 )
            userup1_code.insert(0, 'กรุณากรอกรหัสพนักงาน')
            userup1_code.bind('<FocusIn>',onc_enter)
            userup1_code.bind('<FocusOut>',onc_leave)

            def onn_enter(e):
                userup_name.delete(0,'end')
            def onn_leave(e):
                entry1=userup_name.get()
                if entry1=='':
                    userup_name.insert(0,'กรุณากรอกชื่อ-สกุล')
            userup_name = tk.Entry(popup, width=20,bd=2 ,fg='black', border=0, bg='White', textvariable=name)
            userup_name.place(x=410, y=427)
            userup_name.insert(0, 'กรุณากรอกชื่อ-สกุล')
            userup_name.bind('<FocusIn>',onn_enter)
            userup_name.bind('<FocusOut>',onn_leave)

            choiceage = ['เลือกอายุ'] + [str(i) + ' ปี' for i in range(18, 61)]
            user_age = ttk.Combobox(popup, values=choiceage, textvariable=age)
            user_age.place(x=430, y=467)
            user_age.current(0)

            choicedepart = ['เลือกแผนก','การเงิน','การผลิต','การตลาด','การขาย','บริการ','บุคคล']
            user_depart = ttk.Combobox(popup, values=choicedepart, textvariable=depart)
            user_depart.place(x=450, y=507)
            user_depart.current(0)

            btedit = tk.Button(popup, text='แก้ไข',justify='center', bg='#7e7d7a', fg='black',width=17, border=0,command=update_data)
            btedit.place(x=470, y=545)

        #Guiaccount
        imgac = Image.open(r'Profile.png')
        root.imgac=ImageTk.PhotoImage(imgac)
        acc=Label(popup, image=root.imgac, bg='white',bd=0)
        acc.place(x=273,y=3)

        def on_enter(e):
            user_log.delete(0,'end')
        def on_leave(e):
            entry1=user_log.get()
            if entry1=='':
                user_log.insert(0,'กรุณากรอกรหัสพนักงานเพื่อค้นหา')
        user_log = Entry(popup, width=27, fg='black', border=0, bg='White', textvariable=code)
        user_log.place(x=412, y=59)
        user_log.insert(0, 'กรุณากรอกรหัสพนักงานเพื่อค้นหา')
        user_log.bind('<FocusIn>',on_enter)
        user_log.bind('<FocusOut>',on_leave)
        logi = tk.Button(popup,text='ค้นหา', justify='center', bg='#1bff76',fg='black', bd=0, activebackground='#1bff76',command=logcode)
        logi.place(x=605, y=56)
  
        #เเสดงภาพ
        productsshow1 = Label(popup, cursor="heart",bg='gray', bd=5)
        productsshow1.place(x=435, y=118,height=140,width=140)
        productsshow = Label(popup, cursor="heart",bg='white', borderwidth="2px")
        productsshow.place(x=437, y=120,height=135,width=135)

        listbox_code = Listbox(popup, width=30, height=0)
        listbox_code.place(x=669, y=130)
        listbox_name = Listbox(popup, width=30, height=0)
        listbox_name.place(x=669, y=170)
        listbox_age = Listbox(popup, width=30, height=0)
        listbox_age.place(x=669, y=210)
        listbox_depart = Listbox(popup, width=30, height=0)
        listbox_depart.place(x=669, y=250)
        Btedit = tk.Button(popup, text='Edit', justify='center', bg='gray', bd=0, activebackground='#7e7d7a',command=edit)
        Btedit.place(x=956, y=94, width=42, height=20)

        listbox_record = Listbox(popup, width=58, height=10)
        listbox_record.place(x=659, y=360)
        xsb = Scrollbar(popup, orient=HORIZONTAL, command=listbox_record.xview)
        ysb = Scrollbar(popup, orient=VERTICAL, command=listbox_record.yview)
        listbox_record.config(xscrollcommand=xsb.set, yscrollcommand=ysb.set)
        xsb.place(x=659, y=520,width=359)
        ysb.place(x=1000, y=360,height=167)
        
        listbox_cal = Listbox(popup, width=15, height=0,bd=2)
        listbox_cal.place(x=770, y=540)
        Btcal = tk.Button(popup,text='รายได้ทั้งหมด', justify='center', bg='#7e7d7a',fg='black', bd=0, activebackground='#7e7d7a',command=calculatedisplay)
        Btcal.place(x=870, y=540, width=60, height=20) 
    def in_out():
        login_times = {}
        logout_times = {}
        code=StringVar()
        def logcodein():
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            find_user = ('SELECT * FROM account WHERE code =?')
            cursor.execute(find_user, [user_log.get()])
            result = cursor.fetchone()
            current_time = datetime.datetime.now()
            selected_value = user_io.get()
            income = ''
            
            if result:
                code,name, age, department= result
                is_late = current_time.time() > datetime.time(10, 0)
                late_message = "เข้างานช้า(late)" if is_late else ""
                listbox_show.insert(END, f"เข้างาน  |  ชื่อ-สกุล: {name}, เเผนก: {department},เวลา: {current_time}{late_message}")
                login_times[code] = current_time
                
                conn.close()
                conn_new = sqlite3.connect(r'mydatabaseinout.db') 
                cursor_new = conn_new.cursor()
                cursor_new.execute('INSERT INTO record (inout,code, name, depart, time, late ,income) VALUES (?, ?, ?, ?, ?,?,?)',(selected_value,code, name, department, current_time,late_message,income))
                conn_new.commit()
                conn_new.close()

                user_log.delete(0, 'end')
                user_io.delete(0,'end')
            else:
                conn.close()
                popup.grab_set()
                messagebox.showerror('Error', 'ไม่พบข้อมูลบัญชี')
                popup.grab_release()
        def logcodeout():
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            find_user = ('SELECT * FROM account WHERE code =?')
            cursor.execute(find_user, [user_log.get()])
            result = cursor.fetchone()
            current_time = datetime.datetime.now()
            selected_value = user_io.get()
            income=''
           
            if result:
                code,name, age, department= result
                is_late = current_time.time() > datetime.time(10, 0)
                late_message = " (เข้างานช้า)" if is_late else ""
                listbox1_show.insert(END, f"เข้างาน  |  ชื่อ-สกุล: {name}, เเผนก: {department},เวลา: {current_time}")
                # Store logout time for the user
                logout_times[code] = current_time
                conn.close()
                conn_new = sqlite3.connect(r'mydatabaseinout.db') 
                cursor_new = conn_new.cursor()
                late_message=''
                cursor_new.execute('INSERT INTO record (inout,code, name, depart, time, late ,income) VALUES (?, ?, ?, ?, ?,?,?)',(selected_value,code, name, department, current_time,late_message,income))
                conn_new.commit()
                conn_new.close()
   
                user_log.delete(0, 'end')
                user_io.delete(0,'end')
            else:
                conn.close()
                popup.grab_set()
                messagebox.showerror('Error', 'ไม่พบข้อมูลบัญชี')
                popup.grab_release()
        def getusercode(code):    #defเเสดงชื่อ
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            find_user = ('SELECT name, depart FROM account WHERE code = ?')
            cursor.execute(find_user, [code])
            result = cursor.fetchone()
            conn.close()
            if result:
                name, department = result
                return name, department   
            else:
                return None
        def calculate_income():
            popup.grab_set()
            user_income = {} 
            for code, login_time in login_times.items():
                if code in logout_times:
                    logout_time = logout_times[code]
                    user_info = getusercode(code)  
                    if user_info:
                        name, department = user_info
                        try:
                            hourly_rate = 50  # เรทต่อชม.
                            current_time = datetime.datetime.now()
                            is_late = login_time.time() > datetime.time(9, 0)
                            is_late2 = login_time.time() > datetime.time(13, 0)
                            work_hours = (logout_time - login_time).total_seconds() / 3600
                            income = work_hours * hourly_rate
                            income = round(income, 3)
                            if is_late:        
                                late_hours = login_time.time().hour - 10  # ชั่วโมงที่เข้าสายเกิน 10 โมง
                                late_penalty = late_hours * 10  # หักเงิน -10 บาทต่อชั่วโมงที่เข้าสายเกิน 10 โมง
                                income -= late_penalty
                            if is_late2:
                                income=0
                            user_income[code] = income
                        except ValueError:
                            messagebox.showwarning("ข้อผิดพลาด", f"มีปัญหาในการคำนวณรายได้สำหรับ {name} (รหัส {code})")
                    else:
                        messagebox.showerror('Error', f'ไม่พบข้อมูลผู้ใช้งานรหัส {code}')
            income_message = "รายได้ของแต่ละ user:\n"
            for code, income in user_income.items():
                user_info = getusercode(code) 
                if user_info:
                    current_time = datetime.datetime.now()
                    current_date = current_time.date()
                    name, department = user_info
                    late_message = " (เข้างานช้า)" if is_late else ""
                    income_message += f"วันที่ {current_date}, ชื่อ {name}, แผนก {department}: {income:.2f} บาท{late_message}\n"
                    selected_value=''
                    department=''
                    depart='รายได้'
                    conn_new = sqlite3.connect(r'mydatabaseinout.db')  
                    cursor_new = conn_new.cursor()
                    cursor_new.execute('INSERT INTO record (inout,code, name, depart, time, late , income) VALUES (?, ?, ?, ?, ?, ?,?)',(selected_value,code, name, depart, current_date,late_message,income))
                    listbox_show.delete(0,tk.END)
                    listbox1_show.delete(0,tk.END)
                    conn_new.commit()
                    conn_new.close()
            popup.grab_set()
            messagebox.showinfo("รายได้ของแต่ละ user", income_message)
            popup.grab_release()
        def delete(event=None):
            popup.grab_set()
            selected_item_index = listbox_show.curselection()
            selected_item_index1 = listbox1_show.curselection()
            def deletee():
                if selected_item_index:
                    selected_index = selected_item_index[0]
                    selected_item = listbox_show.get(selected_index)
                    listbox_show.delete(selected_index)
                    delete_selected_record(selected_item)
                elif selected_item_index1:
                    selected_index = selected_item_index1[0]
                    selected_item = listbox1_show.get(selected_index)
                    listbox1_show.delete(selected_index)
                    delete_selected_record(selected_item)
                else:
                    popup.grab_set()
                    messagebox.showinfo('Info', 'กรุณาเลือกรายการที่ต้องการลบ.')
                    popup.grab_release()
            def delete_selected_record(selected_item):
                try:
                    conn = sqlite3.connect(r'mydatabaseinout.db')
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM record WHERE id = (SELECT MAX(id) FROM record);")
                    conn.commit()
                    conn.close()
                    popup.grab_set()
                    messagebox.showinfo("Success",'ลบข้อมูลเเล้ว')
                    popup.grab_release()
                    btdelete.destroy()
                except sqlite3.Error as e:
                    popup.grab_set()
                    messagebox.showerror("Error:", 'ไม่พบข้อมูล')
                    popup.grab_release()
            popup.grab_release()        
            btdelete = Button(popup, width=10, height=0, pady=7, text='ลบ', bg='#7e7d7a', fg='white', border=0, command=deletee)
            btdelete.place(x=805, y=559)
        def time():
            timee=strftime("%I:%M:%S %p")
            time_label.config(text=timee)
            date=strftime("%A %d %B %Y")
            date_label.config(text=date)
            time_label.after(1000, time)# 1000 milliseconds (1 second)
        #Gui in-out
        acc2 = Image.open('homein.png')
        root.acc=ImageTk.PhotoImage(acc2)
        bthome=Label(popup, image=root.acc, bg='white',bd=0)
        bthome.place(x=273,y=3)

        #time
        time_label=Label(popup,font={"Calibri",90},bg='#7e7d7a',fg='white')
        time_label.place(x=450,y=144)       

        date_label=Label(popup,bg='#7e7d7a',fg='white')
        date_label.place(x=432,y=184)    

        def check():
            choiceio = user_io.get()
            if choiceio =='IN-เข้างาน     ':
                logcodein()
            else:
                logcodeout()
        choiceio = ['IN-เข้างาน     ','OUT-ออกงาน']
        user_io=ttk.Combobox(popup,values=choiceio,textvariable=io)
        user_io.place(x=770,y=151)
        user_io.current(0)
        def on_enter(e):
            user_log.delete(0,'end')
        def on_leave(e):
            entry1=user_log.get()
            if entry1=='':
                user_log.insert(0,'กรุณากรอกรหัสพนักงาน')
        user_log = Entry(popup, width=17, fg='black',font=40, border=0, bg='White', textvariable=code)
        user_log.place(x=763, y=189)
        user_log.insert(0, 'กรุณากรอกรหัสพนักงาน')
        user_log.bind('<FocusIn>',on_enter)
        user_log.bind('<FocusOut>',on_leave)
        logi=Button(popup,text='บันทึก',bg='#7e7d7a', justify='center',fg='white',activebackground='#7e7d7a',border=0,command=check)
        logi.place(x=822, y=225)
        btcal=Button(popup,text='คำนวณเงินรายวัน', justify='center',bg='#7e7d7a',fg='white',activebackground='#7e7d7a',border=0,command=calculate_income)
        btcal.place(x=800, y=260)
        
        #listbox in
        listbox_show = tk.Listbox(popup, width=55, height=3)
        listbox_show.place(x=673, y=380)
        listbox_show.bind('<<ListboxSelect>>',delete)
        xsb = Scrollbar(popup, orient=HORIZONTAL, command=listbox_show.xview)
        listbox_show.config(xscrollcommand=xsb)
        xsb.place(x=673, y=420,width=335)
        #listbox out
        listbox1_show = tk.Listbox(popup, width=55, height=3)
        listbox1_show.place(x=673, y=476)
        listbox1_show.bind('<<ListboxSelect>>', delete)
        xsb1 = Scrollbar(popup, orient=HORIZONTAL, command=listbox1_show.xview)
        listbox1_show.config(xscrollcommand=xsb)
        xsb1.place(x=673, y=516,width=335)

        time()
    def leavework():
        def send_emailleave():
            user_top=user_topic.get()
            depart_input1 = listbox_depart.get(0, tk.END)
            name_input1 = listbox_name.get(0, tk.END)
            email_secder = "workwipada155@gmail.com"
            email_password = "fawc uzfs uder mlxa"
            email_recive = "solomonthailand02@gmail.com"
            subject = "ลางาน"
            body = f"{name_input1}\n{depart_input1}\nเเจ้งเรื่อง:{user_top}"
            listbox_name.delete(0,tk.END),listbox_depart.delete(0,tk.END),user_topic.delete(0,tk.END)
            upload_show.config(text="")
            popup.grab_set()
            messagebox.showinfo("ลางาน","เเจ้งเรื่องเรียบร้อยเเล้ว")
            popup.grab_release()
            # ใส่ parameters ไปที่ emailMessage
            em = EmailMessage()
            em  ['From'] = email_secder
            em  ['To'] = email_recive
            em ['Subject'] = subject
     
            em.set_content (body)
    
            context = ssl.create_default_context()
            if global_photo_file_path:
                with open(global_photo_file_path, 'rb') as file:
                    file_data = file.read()
                    em.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=global_photo_file_path)
            #connect email
            with smtplib.SMTP_SSL ('smtp.gmail.com',465,context=context) as smtp:
            # login and send email
                smtp.login(email_secder,email_password)
                smtp.sendmail(email_secder,email_recive,em.as_string())
        def open_file_dialog():
            global global_photo_file_path
            file_path = filedialog.askopenfilename(filetypes=[("Image files, PDF files", "*.jpg *.png *.gif *.pdf")])
            if file_path:
                global_photo_file_path = file_path
                for widget in popup.winfo_children():
                    if isinstance(widget, Label) and widget.cget("text").startswith("Selected File:"):
                        widget.destroy()
                display_file(file_path)
                popup.grab_set()
                upload_show.config(text="File: {}".format(file_path))
                popup.grab_release()
        def display_file(file_path):
            global global_photo, selected_file_label, global_photo_file_path
            if file_path.lower().endswith(('.jpg', '.png', '.gif')): #ตรวจสอบ file_path
                image = Image.open(file_path)
                global_photo = ImageTk.PhotoImage(image)
            elif file_path.lower().endswith('.pdf'):
                pdf = fitz.open(file_path)
                pdf_page = pdf.load_page(0)  # โหลดหน้าเเรกpdf
                # แปลงหน้า PDF เป็นรูปภาพ
                image = pdf_page.get_pixmap()
                image = Image.frombytes("RGB", [image.width, image.height], image.samples)
        def list1():
            search_clicked = 0
            click = leave_use.get()  
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            try:
                find_user1 = ('SELECT name, depart FROM account WHERE code = ?')
                cursor.execute(find_user1, [click])
                result_list = cursor.fetchone()
                if result_list:
                    name, department = result_list
                    if  search_clicked<1:
                        listbox_name.delete(0, tk.END)
                        listbox_depart.delete(0, tk.END)
                    listbox_name.insert(tk.END, f"ชื่อ-สกุล: {name}")
                    listbox_depart.insert(tk.END, f"แผนก: {department}")
                    leave_use.delete(0,tk.END)
                    search_clicked += 1
                else:
                    popup.grab_set()
                    messagebox.showerror('Error', 'ไม่พบข้อมูลบัญชี')
                    popup.grab_release()
            except sqlite3.Error as e:
                popup.grab_set()
                messagebox.showerror('Error', 'เกิดข้อผิดพลาดในการอ่านข้อมูลบัญชี: ' + str(e))
                popup.grab_release()
        
        #Guileavework
        acc3 = Image.open('homeleave.png')
        root.acc3=ImageTk.PhotoImage(acc3)
        lepage=Label(popup, image=root.acc3, bg='white',bd=0)
        lepage.place(x=273,y=3)

        upload_show = Label(popup,text="",bg='#fcfcfc',fg='black',font=("Helvetica", 7,"bold"))
        upload_show.place(x=650,y=455)

        def on_enter(e):
            leave_use.delete(0,'end')
        def on_leave(e):
            leave=leave_use.get()
            if leave=='':
                leave_use.insert(0,'รหัสพนักงาน')
        leave_use = Entry(popup, width=10,bd=0,bg='#fcfcfc',font=30)
        leave_use.place(x=680,y=200)
        leave_use.insert(0, 'รหัสพนักงาน')
        leave_use.bind('<FocusIn>',on_enter)
        leave_use.bind('<FocusOut>',on_leave)
        submit = Button(popup, text="ตกกลง",bg='#7e7d7a',justify='center',fg='white',bd=0,activebackground='#7e7d7a',command=list1)
        submit.place(x=704,y=238)

        listbox_name = Listbox(popup, width=25, height=0)
        listbox_name.place(x=650,y=315)
        listbox_depart = Listbox(popup, width=25, height=0)
        listbox_depart.place(x=650,y=375)
        choicetopic = ['เเจ้งเรื่อง','ลาป่วย','ลาพักร้อน','ลาคลอดบุตร']
        user_topic=ttk.Combobox(popup,values=choicetopic) #,textvariable=depart
        user_topic.place(x=650,y=435)
        user_topic.current(0)

        upload_leave = Button(popup,text="อัปโหลดไฟล์", command=open_file_dialog)
        upload_leave.place(x=800,y=432)

        submit1 = Button(popup, text="ส่ง",command=send_emailleave,bg='#7e7d7a',bd=0,fg='white',activebackground='#7e7d7a',width=10)
        submit1.place(x=682,y=478)
        
    #GUI home
    popup= Toplevel(root)
    popup.title("Solomon Thailand Ltd.")
    width = 1080
    height = 680
    screenwidth = popup.winfo_screenwidth()
    screenheight = popup.winfo_screenheight()
    x = (screenwidth - width) // 2
    y = (screenheight - height) // 2
    alignstr = f"{width}x{height}+{x}+{y}"
    popup.geometry(alignstr)
    popup.resizable(False, False)
    #ใส่เป็นพื้นหลัง
    img = Image.open('mycheck.png')
    root.img=ImageTk.PhotoImage(img)
    pop=Label(popup, image=root.img, bg='white')
    pop.place(x=0,y=0)
 
    def check():
        popup.destroy()
        info_home()
    slm = Image.open('ml.png')
    root.slm=ImageTk.PhotoImage(slm)
    Button_523=tk.Button(popup,image=root.slm,bg='blue',justify='center',bd=0,activebackground='#6aa2ec',command=check)
    Button_523.place(x=29,y=18,width=193,height=59)

    ac = Image.open('myac.png')
    ac1 = Image.open('myac1.png')
    root.ac = ImageTk.PhotoImage(ac)
    root.ac1 = ImageTk.PhotoImage(ac1)
    def onac_enter(e):
        Btacount.config(image=root.ac1)
    def onac_leave(e):
        Btacount.config(image=root.ac)
    Btacount = tk.Button(popup, image=root.ac, justify='center', bg='gray', bd=0, activebackground='#6aa2ec',command=account)
    Btacount.place(x=39, y=193, width=172 , height=46 )
    Btacount.bind('<Enter>', onac_enter)
    Btacount.bind('<Leave>', onac_leave)

    io = Image.open('myio.png')
    io1 = Image.open('myio1.png')
    root.io = ImageTk.PhotoImage(io)
    root.io1 = ImageTk.PhotoImage(io1)
    def onio_enter(e):
        btin_out.config(image=root.io1)
    def onio_leave(e):
        btin_out.config(image=root.io)
    btin_out=tk.Button(popup,image=root.io,justify='center',bg='gray',bd=0,activebackground='#6aa2ec',command=in_out)
    btin_out.place(x=40,y=279,width=163,height=55)
    btin_out.bind('<Enter>', onio_enter)
    btin_out.bind('<Leave>', onio_leave)

    leave = Image.open('myleave.png')
    leave1 = Image.open('myleave1.png')
    root.lv = ImageTk.PhotoImage(leave)
    root.lv1 = ImageTk.PhotoImage(leave1)
    def onlv_enter(e):
        btleave.config(image=root.lv1)
    def onlv_leave(e):
        btleave.config(image=root.lv)
    btleave=tk.Button(popup,image=root.lv,justify='center',bg='gray',bd=0,activebackground='#6aa2ec',command=leavework)
    btleave.place(x=25 ,y=365,width=181,height=61)
    btleave.bind('<Enter>', onlv_enter)
    btleave.bind('<Leave>', onlv_leave) 

    exit1 = Image.open('Back.png')
    exit2 = Image.open('Back1.png')
    root.ex = ImageTk.PhotoImage(exit1)
    root.ex1 = ImageTk.PhotoImage(exit2)
    def onex_enter(e):
        btexit.config(image=root.ex1)
    def onex_leave(e):
        btexit.config(image=root.ex)
    def back():
        popup.grab_set()
        confirmation = messagebox.askquestion("ยืนยัน", "คุณต้องการออกจากระบบหรือไม่?")
        popup.grab_release()
        if confirmation == "yes":
            popup.destroy()
        else:
            pass
    btexit=tk.Button(popup,image=root.ex,justify='center',bg='gray',bd=0,activebackground='#6aa2ec',command=back)
    btexit.place(x=2,y=570,width=75,height=65)
    btexit.bind('<Enter>', onex_enter)
    btexit.bind('<Leave>', onex_leave)
def signin():
    code=StringVar()
    name=StringVar()
    age=StringVar()
    depart=StringVar()
    def signin():
        u_code=user_code.get()
        u_name=user_name.get()
        u_age=user_age.get()
        u_depart=user_depart.get()
        if len(u_code) == 6 :
            if user_code.get()==''or user_name.get()=='' or user_age.get()=='' or user_depart.get()=='':
                messagebox.showerror('Error','กรุณากรอกข้อมูลให้ครบ')
            else:
                try:
                    conn = sqlite3.connect(r'mydatabaseinout.db')
                    cursor = conn.cursor()
                    cursor.execute('INSERT INTO account (code,name, age, depart) VALUES (?, ?, ?,?)', (u_code,u_name,u_age,u_depart))
                    conn.commit()
                    conn.close()
                    clear()
                    messagebox.showinfo('Success','กรอกข้อมูลเรียบร้อย')

                    file_label.config(text="")
                    file_pic = None
                except Exception as es:
                    messagebox.showerror('Error','กรุณากรอกข้อมูลใหม่')
        else:
            messagebox.showerror('Error', 'รหัสพนักงาน 6 ตัว')          
    def add():
        code = user_code.get()
        file_pic = filedialog.askopenfilename()
        if code and file_pic:
            try:
                conn = sqlite3.connect(r'mydatabaseinout.db')
                cursor = conn.cursor()
                with open(file_pic, 'rb') as file:
                    picture = file.read()
                    file_label.config(text="File: {}".format(file_pic))
                cursor.execute("INSERT INTO img (code, data) VALUES (?, ?)", (code, picture))
                conn.commit()
                conn.close()

            except Exception as es:
                messagebox.showerror('Error', 'เกิดข้อผิดพลาดขณะเพิ่มรูปภาพ')
        else:
            messagebox.showerror('Error', 'โปรดระบุรหัสและเลือกไฟล์ภาพ')
    def clear():
        code.set('')
        name.set('')
        age.set('')
        depart.set('')
    #GUI signin
    pop_signin=root
    
    sign = Image.open('Homesign.png')
    root.sign=ImageTk.PhotoImage(sign)
    sig=Label(pop_signin, image=root.sign, bg='white')
    sig.place(x=0)
    
    def on_enter(e):
        user_code.delete(0,'end')
    def on_leave(e):
        name=user_code.get()
        if name=='':
            user_code.insert(0,'กรุณากรอกรหัสพนักงาน')
    user_code=Entry(pop_signin,width=30,fg='black',font=70,border=0,bg='white',textvariable=code)
    user_code.place(x=395,y=162)
    user_code.insert(0,'กรุณากรอกรหัสพนักงาน')
    user_code.bind('<FocusIn>',on_enter)
    user_code.bind('<FocusOut>',on_leave)
    
    def on_enter(e):
        user_name.delete(0,'end')
    def on_leave(e):
        name=user_name.get()
        if name=='':
            user_name.insert(0,'กรุณากรอกชื่อ-สกุล')
    user_name=Entry(pop_signin,width=30,fg='black',font=70,border=0,bg='white',textvariable=name)
    user_name.place(x=395,y=244)
    user_name.insert(0,'กรุณากรอกชื่อ-สกุล')
    user_name.bind('<FocusIn>',on_enter)
    user_name.bind('<FocusOut>',on_leave)

    choiceage =['เลือกอายุ']+[str(i)+ ' ปี' for i in range(18, 61)]
    user_age=ttk.Combobox(pop_signin,values=choiceage,font=70,textvariable=age)
    user_age.place(x=390,y=335)
    user_age.current(0)

    choicedepart = ['เลือกเเผนก','การเงิน','การผลิต','การตลาด','การขาย','บริการ','บุคคล']
    user_depart=ttk.Combobox(pop_signin,values=choicedepart,font=70,textvariable=depart)
    user_depart.place(x=390,y=425.5)
    user_depart.current(0)
    upload_leave = Button(pop_signin,text="อัปโหลดไฟล์",command=add)
    upload_leave.place(x=620,y=424)
    
    file_label = tk.Label(pop_signin, text="",bg='white')
    file_label.place(x=388,y=450)

    sign_in=Button(pop_signin,width=39,pady=7,text='SIGN-IN',font=70,bg='#313030',fg='white',border=0,command=signin)
    sign_in.place(x=390,y=517)
    def backl():
        user_code.destroy()
        user_name.destroy()
        user_age.destroy()
        user_depart.destroy()
        upload_leave.destroy()
        file_label.destroy()
        sign_in.destroy
        backk.destroy()
        sig.destroy()
        sign_in.destroy()
    backk=Button(pop_signin,text='<-BACK',bg='#7e7d7a',activebackground='#7e7d7a',fg='white',command=backl)
    backk.place(x=1000,y=30)  
def login():
    code=StringVar()
    io=StringVar()
    def logcode():
        conn = sqlite3.connect(r'mydatabaseinout.db')
        cursor = conn.cursor()
        find_user=('SELECT * FROM account WHERE code =?')
        cursor.execute(find_user,[user_log.get()])
        result= cursor.fetchall()
        user=user_log.get()
        if user == '1234' :
            messagebox.showinfo('ADMIN','ล็อกอินหน้าADMINสำเร็จ')
            adminb()
        elif result:
            messagebox.showinfo('Success','ล็อกอินสำเร็จ')
            info_home()
        else:
            messagebox.showerror('Fail','รหัสพนักงานไม่ถูกต้อง')
    #GUI signin
    pop_login=root
    #ใส่เป็นพื้นหลัง
    log = Image.open('Homelogin.png')
    root.log=ImageTk.PhotoImage(log)
    page=Label(pop_login, image=root.log, bg='white')
    page.place(x=0)

    def on_enter(e):
        user_log.delete(0,'end')
    def on_leave(e):
        name=user_log.get()
        if name=='':
            user_log.insert(0,'กรุณากรอกรหัสพนักงาน')
    user_log=Entry(pop_login,width=45,fg='black',font=120,border=0,bg='White',textvariable=code)
    user_log.place(x=355,y=315)
    user_log.insert(0,'กรุณากรอกรหัสพนักงาน')
    user_log.bind('<FocusIn>',on_enter)
    user_log.bind('<FocusOut>',on_leave)
    regis=Button(pop_login,width=17,text="Don't have an account?",font=5,bg='#f6f6f6',fg='black',border=0,activebackground='#f6f6f6',command=signin)
    regis.place(x=300,y=470)
    logi=Button(pop_login,width=10,text='เข้าสู่ระบบ',font=70,bg='#313030',fg='white',activebackground='#313030',border=0,command=logcode)
    logi.place(x=700,y=470)
   
    def back():
        page.destroy()
        user_log.destroy()
        regis.destroy()
        logi.destroy()
        backl.destroy()
    backl=Button(pop_login,text='<-BACK',bg='#7e7d7a',activebackground='#7e7d7a',fg='white',command=back)
    backl.place(x=1000,y=30)
def about():
    about_window = Toplevel(root)
    about_window.title("Solomon Thailand Ltd.")
    width = 1080
    height = 680
    screenwidth = about_window.winfo_screenwidth()
    screenheight = about_window.winfo_screenheight()
    x = (screenwidth - width) // 2
    y = (screenheight - height) // 2
    alignstr = f"{width}x{height}+{x}+{y}"
    about_window.geometry(alignstr)
    about_window.resizable(False, False)
    
    #สร้างเฟรมเนื้อหา
    content_frame = ttk.Frame(about_window)
    content_frame.pack(fill="both", expand=True)
    # Add scrollbar
    y_scrollbar = ttk.Scrollbar(content_frame, orient="vertical")
    y_scrollbar.pack(side="right", fill="y")

    frame=Frame(content_frame,width=1080,height=80,bg="#313030")
    frame.pack()
   
    canvas = tk.Canvas(content_frame, yscrollcommand=y_scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    # Link scrollbar
    y_scrollbar.config(command=canvas.yview)
    # Create  content 
    scrollable_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    img1 = Image.open("pl.png")
    img1 = ImageTk.PhotoImage(img1)
    label = ttk.Label(scrollable_frame, image=img1, background='white')
    label.image = img1
    label.pack(pady=2)
 
    img_height = img1.height()
    current_y = img_height
    
    img = Image.open("pl1.png")
    img = ImageTk.PhotoImage(img)
    img_label = ttk.Label(scrollable_frame, image=img, background='white')
    img_label.image = img
    img_label.pack(pady=2)
    # Calculate the position for the next image
    current_y += img_height
    current_y += img_height
    img2 = Image.open("thank.png")
    img2 = ImageTk.PhotoImage(img2)
    img2_label = ttk.Label(scrollable_frame, image=img2, background='white')
    img2_label.image = img2
    img2_label.pack()
    # อัปเดตCanvas Scroll ให้พอดีกับเนื้อหาทั้งหมด
    canvas.config(scrollregion=(0, 0, canvas.winfo_width(), current_y))
    canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

    def home():
        about_window.withdraw()
    slmac= Image.open('MONTHAILAND.png')
    root.slmac = ImageTk.PhotoImage(slmac)
    slm=tk.Button(about_window,image=root.slmac,bg='blue',justify='center',bd=0,activebackground='#313030',command=home).place(x=12 ,y=14,width=190,height=59)

    #about
    btraac = Image.open('Abb.png')
    btr1aac = Image.open('Abb1.png')
    root.btraac = ImageTk.PhotoImage(btraac)
    root.btr1aac = ImageTk.PhotoImage(btr1aac)
    def onbtr_enter(e):
        Btraac.config(image=root.btr1aac)
    def onbtr_leave(e):
        Btraac.config(image=root.btraac)
    def ab():
        about_window.destroy()
        about()
    Btraac = tk.Button(about_window,image=root.btraac,bg='blue',justify='center',bd=0,activebackground='#313030',command=ab)
    Btraac.place(x=375,y=32, width=98, height=33)
    Btraac.bind('<Enter>', onbtr_enter)
    Btraac.bind('<Leave>', onbtr_leave)

    #logIn
    btlogac = Image.open('Lgg.png')
    btlog13ac = Image.open('Lgg1.png')
    root.btlogac = ImageTk.PhotoImage(btlogac)
    root.btlog13ac = ImageTk.PhotoImage(btlog13ac)
    def onbtlog_enter(e):
        Btloginac.config(image=root.btlog13ac)
    def onbtlog_leave(e):
        Btloginac.config(image=root.btlogac)
    def logi():
        about_window.destroy()
        login()
    Btloginac = tk.Button(about_window,image=root.btlogac,bg='blue',justify='center',bd=0,activebackground='#313030',command=logi)
    Btloginac.place(x=500,y=30, width=98, height=33)
    Btloginac.bind('<Enter>', onbtlog_enter)
    Btloginac.bind('<Leave>', onbtlog_leave)

    #contact
    btrac = Image.open('CONTACT.png')
    btr1ac = Image.open('CONTACT1.png')
    root.btrac = ImageTk.PhotoImage(btrac)
    root.btr1ac = ImageTk.PhotoImage(btr1ac)
    def onbtr_enter(e):
        Btrac.config(image=root.btr1ac)
    def onbtr_leave(e):
        Btrac.config(image=root.btrac)

    def conta():
        about_window.destroy()
        contack()
    Btrac = tk.Button(about_window,image=root.btrac,bg='blue',justify='center',bd=0,activebackground='#313030',command=conta)
    Btrac.place(x=625,y=27, width=138, height=38)
    Btrac.bind('<Enter>', onbtr_enter)
    Btrac.bind('<Leave>', onbtr_leave)

    #signin
    bt1ac = Image.open('Sign1.png')
    btac = Image.open('Sign2.png')
    root.bt1ac = ImageTk.PhotoImage(bt1ac)
    root.btac = ImageTk.PhotoImage(btac)
    def onbt_enter(e):
        Btslogac.config(image=root.btac)
    def onbt_leave(e):
        Btslogac.config(image=root.bt1ac)
    def sign():
        about_window.destroy()
        signin()
    Btslogac = tk.Button(about_window,image=root.bt1ac,bg='blue',justify='center',bd=0,activebackground='#313030',command=sign)
    Btslogac.place(x=939.97,y=22.81, width=117.21, height=39.38)
    Btslogac.bind('<Enter>', onbt_enter)
    Btslogac.bind('<Leave>', onbt_leave)  
def adminb():
    def create_tab(tab_control, tab_title):
        frame = ttk.Frame(tab_control)
        tab_control.add(frame, text=tab_title)
        return frame
    def add_account(tab_frame, label_text):
        def fetch_data():
            filter_text = filter_entry.get()
            if filter_text:
                conn = sqlite3.connect(r'mydatabaseinout.db')
                cursor = conn.cursor()
                query = "SELECT * FROM account WHERE code = ?"
                cursor.execute(query, (filter_text,))
                data = cursor.fetchall()
                conn.close()
                listbox.delete(0, tk.END)
                for row in data:
                    # ดึงข้อมูลจากแต่ละคอลัมน์ในแถว
                    code, name, age, depart = row
                    listbox.insert(tk.END, f"{code} | ชื่อ-สกุล: {name}, อายุ: {age}, แผนก: {depart}")
            else:
                conn = sqlite3.connect(r'mydatabaseinout.db')
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM account")
                data = cursor.fetchall()
                conn.close()
                listbox.delete(0, tk.END)
                for row in data:
                    code, name, age, depart = row
                    listbox.insert(tk.END, f"{code} | ชื่อ-สกุล: {name}, อายุ: {age}, แผนก: {depart}")
            filter_entry.delete(0,END)
        ac = Image.open('Adminac.png')
        root.ac=ImageTk.PhotoImage(ac)
        account=Label(tab_frame, image=root.ac, bg='white')
        account.place(x=0)
        # Entry สำหรับกรองข้อมูล
        filter_entry = ttk.Entry(tab_frame)
        filter_entry.place(x=440,y=85)
        fetch_button = ttk.Button(tab_frame, text="ค้นหา", command=fetch_data)
        fetch_button.place(x=570,y=82)
        listbox = tk.Listbox(tab_frame, width=74, height=27,font=60,bd=0)
        listbox.place(x=370, y=120)
        ysb = Scrollbar(tab_frame, orient=VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=ysb.set)
        ysb.place(x=1080, y=80,height=600)
        fetch_data()
    def add_record(tab_frame, label_text):
        def fetch_data1():
            filter1_text = filter1_entry.get()
            
            if filter1_text:
                conn = sqlite3.connect(r'mydatabaseinout.db')
                cursor = conn.cursor()
                query = "SELECT * FROM record WHERE code = ?"
                cursor.execute(query, (filter1_text,))
                data1 = cursor.fetchall()
                conn.close()
                
                # ลบข้อมูลเก่าออกจาก listbox1
                listbox1.delete(0, tk.END)

                late_count = 0
                fpdfrec = FPDF()
                fpdfrec.add_page()
                fpdfrec.set_font("Arial", size=20,style='B')
                formatted_datarec = []
                fpdfrec.set_text_color(0, 0, 255) 
                rec = 'Record data'
                item_latin = rec.encode("latin-1", "ignore").decode("latin-1")
                fpdfrec.text(90, 7, txt=item_latin)
                if rec:
                    fpdfrec.set_text_color(0, 0, 0) 
                    fpdfrec.set_font("Arial", size=10)
                else:
                    fpdfrec.set_text_color(0, 0, 255) 
                
                formatted_datarec.append(f"")
                formatted_datarec.append(f"IN-OUT-----------------------------------------------------------------------------")
                for row in data1:
                    id, inout, code, name, depart, time, late, income = row
                    if inout and code and name and depart and time:
                        # แสดงข้อมูลใน listbox1
                        listbox1.insert(tk.END, f"{inout}       {code}             {name}           {depart}                       {time}{late}")
                        formatted_datarec.append(f"{inout}|  Code:{code} | Name:{name} | Depart:{depart} | Time:{time}{late}")

                    elif income and time:
                        listbox1.insert(tk.END, f"รายได้ประจำวันของ {name} วันที่ {time}= {income} บาท")
                        formatted_datarec.append( f"income|Name: {name} Time: {time} = {income} bath")
                for row in data1:
                    if row[6] == "เข้างานช้า(late)":
                        late_count += 1
                    elif late_count>=5 :
                        messagebox.showinfo('','คุณถูกพิจารณา')
                formatted_datarec.append(f"")
                formatted_datarec.append(f"LATE--------------------------------------------------------------------------------")
                listbox1.insert(tk.END, f"จำนวนครั้งเข้าสาย| {name} จำนวน {late_count}")
                formatted_datarec.append(f"Late| Name: {name} : {late_count}")
            
                # ลูปเพื่อแทรกข้อมูลลงในไฟล์ PDF
                for item in formatted_datarec:
                    # แปลงข้อความให้เป็น latin-1 ก่อน
                    item_latin1 = item.encode("latin-1", "ignore").decode("latin-1")
                    fpdfrec.multi_cell(0, 5, txt=item_latin1, align="L")

                fpdfrec.output('Rec.pdf')
                display_image()
            filter1_entry.delete(0,END)
        def display_image():
            conn = sqlite3.connect(r'mydatabaseinout.db')
            cursor = conn.cursor()
            user_code = filter1_entry.get()
            try:
                cursor.execute("SELECT data FROM img WHERE code=?", (user_code,))
                pic = cursor.fetchone()
                if pic and pic[0]:
                    try:
                        search_clicked = 0
                        if search_clicked <1:
                            for widget in productsshow.winfo_children():
                                widget.destroy()
                        for  x in enumerate(pic):
                            image = Image.open(BytesIO(x[1]))
                            target_width, target_height = 135, 135
                            image = image.resize((target_width, target_height))
                            image = ImageTk.PhotoImage(image)

                            img_label = Label(productsshow, image=image, bg='blue',bd=10)
                            img_label.image = image
                            img_label.pack()
                        search_clicked += 1
                    except Exception as e:
                        messagebox.showerror(f"Error displaying image: {str(e)}")
                else:
                    messagebox.showerror('Error',"ไม่มีข้อมูล")
            except sqlite3.Error as e:
                messagebox.showerror('Error', 'เกิดข้อผิดพลาดในการอ่านข้อมูลบัญชี: ' + str(e))
        def expenrec():
            import webbrowser
            pdf_url = 'rec.pdf'
            webbrowser.open(pdf_url, new=2)
      
        rec = Image.open('Adminrec.png')
        root.rec=ImageTk.PhotoImage(rec)
        record=Label(tab_frame, image=root.rec, bg='white')
        record.place(x=0)

        # เเสดงรูป
        productsshow = Label(tab_frame, cursor="heart",bg='white', borderwidth="2px")
        productsshow.place(x=100, y=30,height=135,width=135)
        # Entry สำหรับกรองข้อมูล
        filter1_entry = ttk.Entry(tab_frame)
        filter1_entry.place(x=440,y=85)
        fetch1_button = ttk.Button(tab_frame, text="ค้นหา", command=fetch_data1)
        fetch1_button.place(x=570,y=82)
        pdf = Button(tab_frame, text="Download file.",bg='white',bd=0,font=9,activebackground='white',activeforeground='red',fg='blue')
        pdf.place(x=680,y=32)

        printt = Image.open('print.png')
        root.printt=ImageTk.PhotoImage(printt)
        pdfrec = Button(tab_frame, text="Print",image=root.printt, compound="right", command=expenrec)
        pdfrec.place(x=980,y=80)

        listbox1 = tk.Listbox(tab_frame, width=74, height=27, font=("Arial", 12), bd=0)
        listbox1.place(x=370, y=170)
        ysb = Scrollbar(tab_frame, orient=VERTICAL, command=listbox1.yview)
        listbox1.config(yscrollcommand=ysb.set)
        ysb.place(x=1080, y=170, height=400) 
        fetch_data1()
    # สร้างpdf
    # def pdf():
    #     pdfname="รายจ่าย.pdf"
    #     c = canvas.Canvas(pdfname,pagesize=letter)
    #     c.drawString(100,750,"hello")
    #     c.save

    admin=Toplevel(root)
    admin.title("ADMIN | Solomon thailand ltd.")
    width = 1080
    height = 680
    screenwidth = admin.winfo_screenwidth()
    screenheight = admin.winfo_screenheight()
    x = (screenwidth - width) // 2  
    y = (screenheight - height) // 2  
    alignstr = f"{width}x{height}+{x}+{y}"  
    admin.geometry(alignstr)
    admin.resizable(False,False)

    tab_control = ttk.Notebook(admin)
    tab1 = create_tab(tab_control, "รายชื่อ")
    add_account(tab1, "Account")
    tab2 = create_tab(tab_control, "ประวัติ")
    add_record(tab2, "Record")
    tab_control.pack(fill="both", expand=True)
def contack():
    def sendemail():
        fname=fnamec.get()
        lname=lnamec.get()
        maill=mail.get()
        tell=tel.get()
        write=mess.get()
        if fname==" " or lname==" " or maill==" " or write==" " or len(tell)==10:
            email_secder = "workwipada155@gmail.com"
            email_password = "fawc uzfs uder mlxa"
            email_recive = "solomonthailand02@gmail.com"
            subject = "Message"
            body = f"ชื่อ-สกุล:{fname} {lname}\nอีเมล:{maill}\nเบอร์โทรศัพท์:{tell}\n:เรื่อง:{write}"
            clear()
            # ใส่ parameters ไปที่ emailMessage
            em = EmailMessage()
            em  ['From'] = email_secder
            em  ['To'] = email_recive
            em ['Subject'] = subject
            em.set_content (body)
            context = ssl.create_default_context()
            #connect email
            with smtplib.SMTP_SSL ('smtp.gmail.com',465,context=context) as smtp:
            # login and send email
                smtp.login(email_secder,email_password)
                smtp.sendmail(email_secder,email_recive,em.as_string())
        else:
            messagebox.showwarning("Error",'กรุณากรอเบอร์ให้ครบ')
    def clear():
        fnamec.delete(0,END)
        lnamec.delete(0,END)
        mail.delete(0,END)
        tel.delete(0,END)
        mess.delete(0,END)
    contackpage=root
    con = Image.open('D:\python\project\Conta.png')
    root.con=ImageTk.PhotoImage(con)
    contack=Label(contackpage, image=root.con,bd=0 ,bg='white',width=688,height=555)
    contack.place(x=205,y=93)

    fnamec=Entry(contackpage,width=17,fg='black',font=70,border=0,bg='#fcfcfc')
    fnamec.place(x=520,y=255)
    lnamec=Entry(contackpage,width=17,fg='black',font=70,border=0,bg='#fcfcfc')
    lnamec.place(x=707,y=255)
    mail=Entry(contackpage,width=17,fg='black',font=70,border=0,bg='#fcfcfc')
    mail.place(x=520,y=336)
    tel=Entry(contackpage,width=17,fg='black',font=70,border=0,bg='#fcfcfc')
    tel.place(x=707,y=336)
    mess=Entry(contackpage,width=39,bd=1,fg='black',font=70,bg='#fcfcfc')
    mess.place(x=520,y=418)
    send=Button(contackpage,text='Send', justify='center',height=2,width=15,bg='#7e7d7a',fg='white',border=0,command=sendemail)
    send.place(x=520, y=521)
    def back():
        fnamec.destroy()
        lnamec.destroy()
        mail.destroy()
        tel.destroy()
        mess.destroy()
        send.destroy()
        baak.destroy()
        contack.destroy()
    baak=Button(contackpage,text='<-BACK',bg='#7e7d7a',activebackground='#7e7d7a',fg='white',command=back)
    baak.place(x=820,y=130)
#GUI main
root=Tk()
root.title("Solomon Thailand Ltd.")
width = 1080
height = 680
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
x = (screenwidth - width) // 2
y = (screenheight - height) // 2
alignstr = f"{width}x{height}+{x}+{y}"
root.geometry(alignstr)

root.resizable(False, False)

#  contentframe 
content_framehome = ttk.Frame(root)
content_framehome.pack(fill="both", expand=True)

y_scrollbarhome = ttk.Scrollbar(content_framehome, orient="vertical")
y_scrollbarhome.pack(side="right", fill="y")

tab = Image.open("Tab.png")
tab = ImageTk.PhotoImage(tab)
frame=Frame(content_framehome,width=1080,height=100,bg="gray")
frame.pack()
Label(frame,image=tab,width=1080,height=80).pack()

canvas = tk.Canvas(content_framehome, yscrollcommand=y_scrollbarhome.set)
canvas.pack(side="left", fill="both", expand=True)
# Link scrollbar 
y_scrollbarhome.config(command=canvas.yview)
# Create  content
scrollable_framehome = ttk.Frame(canvas)
canvas.create_window((0, 0), window=scrollable_framehome, anchor="nw")

img1 = Image.open("Homepage.png")
img1 = ImageTk.PhotoImage(img1)
label = ttk.Label(scrollable_framehome, image=img1, background='white')
label.image = img1
label.pack()

img_height = img1.height()
current_y = img_height


img = Image.open(r"pl1.png")
img = ImageTk.PhotoImage(img)
img_label = ttk.Label(scrollable_framehome, image=img, background='red')
img_label.image = img
img_label.pack()

canvas.config(scrollregion=(0, 0, canvas.winfo_width(), current_y))
canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

slm= PhotoImage(file='MONTHAILAND.png')
GButton_523=tk.Button(root,image=slm,bg='blue',justify='center',bd=0,activebackground='#313030').place(x=12 ,y=14,width=190,height=59)

#about
btra = Image.open('Abb.png')
btr1a = Image.open('Abb1.png')
root.btra = ImageTk.PhotoImage(btra)
root.btr1a = ImageTk.PhotoImage(btr1a)
def onbtr_enter(e):
    Btr0.config(image=root.btr1a)
def onbtr_leave(e):
    Btr0.config(image=root.btra)
Btr0 = tk.Button(root,image=root.btra,bg='blue',justify='center',bd=0,activebackground='#313030',command=about)
Btr0.place(x=375,y=32, width=98, height=33)
Btr0.bind('<Enter>', onbtr_enter)
Btr0.bind('<Leave>', onbtr_leave)

#contact
btr = Image.open('CONTACT.png')
btr1 = Image.open('CONTACT1.png')
root.btr = ImageTk.PhotoImage(btr)
root.btr1 = ImageTk.PhotoImage(btr1)
def onbtr_enter(e):
    Btr.config(image=root.btr1)
def onbtr_leave(e):
    Btr.config(image=root.btr)
Btr = tk.Button(root,image=root.btr,bg='blue',justify='center',bd=0,activebackground='#313030',command=contack)
Btr.place(x=625,y=27, width=138, height=38)
Btr.bind('<Enter>', onbtr_enter)
Btr.bind('<Leave>', onbtr_leave)

#signin
bt1 = Image.open('Sign1.png')
bt = Image.open('Sign2.png')
root.bt1 = ImageTk.PhotoImage(bt1)
root.bt = ImageTk.PhotoImage(bt)
def onbt_enter(e):
    Btlog.config(image=root.bt)
def onbt_leave(e):
    Btlog.config(image=root.bt1)
Btlog = tk.Button(root,image=root.bt1,bg='blue',justify='center',bd=0,activebackground='#313030',command=signin)
Btlog.place(x=939.97,y=22.81, width=117.21, height=39.38)
Btlog.bind('<Enter>', onbt_enter)
Btlog.bind('<Leave>', onbt_leave)  

#logIn
btlog = Image.open('Lgg.png')
btlog13 = Image.open('Lgg1.png')
root.btlog = ImageTk.PhotoImage(btlog)
root.btlog13 = ImageTk.PhotoImage(btlog13)
def onbtlog_enter(e):
    Btlogin.config(image=root.btlog13)
def onbtlog_leave(e):
    Btlogin.config(image=root.btlog)
Btlogin = tk.Button(root,image=root.btlog,bg='blue',justify='center',bd=0,activebackground='#313030',command=login)
Btlogin.place(x=500,y=30, width=98, height=33)
Btlogin.bind('<Enter>', onbtlog_enter)
Btlogin.bind('<Leave>', onbtlog_leave)

root.mainloop()