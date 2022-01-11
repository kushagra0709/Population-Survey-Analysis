import xlrd
from xlutils.copy import copy
import datetime
from datetime import datetime as ddtt
import matplotlib.pyplot as plt
import random
import mysql.connector as con
import tkinter
from tkinter import messagebox,scrolledtext
import tkinter.ttk as tk
from tabulate import tabulate
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


def open_window():
    global top
    top=tkinter.Toplevel()
    top.geometry("750x750+120+120")
    top.configure(bg="papaya whip")
    but=tk.Button(top,text="close",command=top.destroy)
    but.pack()

def updating():
    l=[]#First name
    m=[]#Surname
    n=[]#Fullname
    o=[]#Age
    p=[]#Dependency
    q=[]#Gender
    r=[]#Number
    s=[]#Address
    t=[]#Email
    status=[]#Status
    book=xlrd.open_workbook("NAME.xlsx")
    first_sheet=book.sheet_by_index(0)
    conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',database="project")
    cur=conn.cursor()
    a=[]
    i=0
    while i!=5:
        b=random.randint(1,200)
        if b not in a:
            a.append(b)
            query="delete from details where SNo="+str(b)
            cur.execute(query)
            z=random.randint(0,429)
            cell=first_sheet.cell(z,0)
            l.append(cell.value.strip())
            cell2=first_sheet.cell(z,1)
            q.append(cell2.value)
            c=random.randint(0,133)
            cell = first_sheet.cell(c,2)
            m.append(cell.value.strip())
            d=random.randint(0,11)
            cell=first_sheet.cell(d,3)
            s.append(cell.value)
            x=l[i]+" "+m[i]
            n.append(x)
            y=random.randint(1,100)
            o.append(y)
            pov = ["below poverty line","middle income","low income","high income"]
            ran = random.randint(0,3)
            dep = pov[ran]
            status.append(dep)
            k=''
            A=9
            for j in range(10):
                k+=str(A)
                A=random.randint(0,9)
            r.append(int(k))
            B=l[i].lower()+str(o[i])+"@gmail.com"
            t.append(B)
            for f in o:
                if f<=20:
                    y="Dependent"
                elif f>=65:
                    y="Dependent"
                else:
                    y="Independent"
            p.append(y)
            i+=1
        elif b in a:
            i-=1
    for i in range(len(a)):
        query=("insert into Details values(%s,%s,%s,%s,%s,%s,%s,%s,%s)")
        tup=(a[i],n[i],o[i],r[i],t[i],s[i],q[i],p[i],status[i])
        cur.execute(query,tup)
    conn.commit()
    cur.close()
    conn.close()
    
book=xlrd.open_workbook("NAME.xlsx")
first_sheet=book.sheet_by_index(0)
cell=first_sheet.cell(0,4)
v=cell.value
excel_date =int(v)
dt = ddtt.fromordinal(ddtt(1900, 1, 1).toordinal() + excel_date - 2).date()
xx=str(dt)
a=datetime.date.today()
now = datetime.datetime.strptime(xx, '%Y-%m-%d').date()
b=a-now
if str(b)!="0:00:00":
    updating()
    print("UPDATING DATABASE...")
    v=str(datetime.datetime.today()).split()[0]
    v1=datetime.datetime.strptime(v, '%Y-%m-%d').date()
    book=xlrd.open_workbook("NAME.xlsx")
    first_sheet=book.sheet_by_index(0)
    wbook=copy(book)
    w_sheet = wbook.get_sheet(0)
    w_sheet.write(0,4,v1)
    wbook.save("NAME.xlsx")  
else:
    pass

def user():
    window1.destroy()
    def sex_ratio():
        conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',database="project")
        cur=conn.cursor()
        query="select gender from details"
        cur.execute(query)
        gender=cur.fetchall()
        men=women=0
        for i in gender:
            if i==('Male',):
                men+=1
            elif i==('Female',):
                women+=1
        conn.commit()
        cur.close()
        conn.close()
        open_window()
        frame4=tkinter.Frame(top,bd=10,bg="papaya whip")
        frame4.pack()
        figure1 = plt.Figure(figsize=(8,4), dpi=90)
        figure1.suptitle("SEX RATIO")
        ax1 = figure1.add_subplot(131)
        ax1.pie([men,women],labels = ["MALE","FEMALE"],explode=[0.1,0],startangle=45,autopct="%1.0f%%")
        
        
        ax2 = figure1.add_subplot(133)
        a2 = FigureCanvasTkAgg(figure1,frame4)
        ax2.bar(1,women,label="Female",color="orange")
        ax2.bar(2,men,label="Male")
        a2.get_tk_widget().pack()
        ax2.legend()
    
    def particulars(x):
        conn = con.connect(user="root",passwd="reckoning",host = "localhost",auth_plugin="mysql_native_password",database="project")
        cur = conn.cursor()
        query = "select * from details where Name like "+"\""+x+"\""
        cur.execute(query)
        a = cur.fetchall()
        
        open_window()
        frame3=tkinter.Frame(top,bd=10,bg="papaya whip")
        frame3.pack(fill="x")
        z=0
        for i in a:
            for j in range(len(i)):
                z+=1
                if j == 1:
                    particular_1=tkinter.Label(frame3,text="NAME---->",font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_1.grid(column=0,row=z)
                    particular_11=tkinter.Label(frame3,text=i[j],font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_11.grid(column=1,row=z)
                elif j == 2:
                    particular_2=tkinter.Label(frame3,text="AGE---->",font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_2.grid(column=0,row=z)
                    particular_12=tkinter.Label(frame3,text=i[j],font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_12.grid(column=1,row=z)
                elif j == 3:
                    particular_3=tkinter.Label(frame3,text="PHONE NUMBER---->",font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_3.grid(column=0,row=z)
                    particular_13=tkinter.Label(frame3,text=i[j],font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_13.grid(column=1,row=z)
                elif j == 4:
                    particular_4=tkinter.Label(frame3,text="EMAIL---->",font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_4.grid(column=0,row=z)
                    particular_14=tkinter.Label(frame3,text=i[j],font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_14.grid(column=1,row=z)
                elif j == 5:
                    particular_5=tkinter.Label(frame3,text="ADDRESS---->",font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_5.grid(column=0,row=z)
                    particular_15=tkinter.Label(frame3,text=i[j],font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_15.grid(column=1,row=z)
                elif j == 6:
                    particular_6=tkinter.Label(frame3,text="GENDER---->",font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_6.grid(column=0,row=z)
                    particular_16=tkinter.Label(frame3,text=i[j],font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_16.grid(column=1,row=z)
                elif j == 7:
                    particular_7=tkinter.Label(frame3,text="DEPENDANCY---->",font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_7.grid(column=0,row=z)
                    particular_17=tkinter.Label(frame3,text=i[j],font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_17.grid(column=1,row=z)
                elif j ==8:
                    particular_8=tkinter.Label(frame3,text="ECONOMIC STATUS---->",font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_8.grid(column=0,row=z)
                    particular_18=tkinter.Label(frame3,text=i[j],font=("Helvetica",15),bg="papaya whip",fg="black")
                    particular_18.grid(column=1,row=z)
            else:
                z+=1
                tkinter.Label(frame3,text="",bg="papaya whip").grid(column=0,row=z)
        conn.commit()
        cur.close()
        conn.close()
    
    def voters_list():
        conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',db="Project")
        cur=conn.cursor()
        Age1=("select Name,Age,Address,Phone_No from details where Age>=18 order by Age")
        cur.execute(Age1)
        Vlist=cur.fetchall()
        count=len(Vlist)
        headers=["Name","Age","Address","Phone Number"]
        open_window()
        frame3=tkinter.Frame(top,bd=10,bg="black")
        frame3.pack(fill="x")
        scrolling_list=scrolledtext.ScrolledText(top,width=600,height=600)
        scrolling_list.pack()
        scrolling_list.insert(tkinter.INSERT,tabulate(Vlist,headers=headers))
        scrolling_list.insert(tkinter.INSERT,"\n_ _ _ _ _ _ _ _ _ _\n")
        scrolling_list.insert(tkinter.INSERT,("Total Count="+str(count)))
        scrolling_list.config(state=tkinter.DISABLED)
    
    def economicdatagraph():
        conn = con.connect(user="root",passwd="reckoning",host = "localhost",auth_plugin="mysql_native_password",database="project")
        cur = conn.cursor()
        cur.execute("select economic_status from details")
        a = cur.fetchall()
        cp = 0 #count bpl
        cl = 0 #count low income
        cm = 0 #count middle income
        ch = 0 #count high income
        for i in a:
            for j in i:
                if j == "below poverty line":
                    cp += 1
                elif j =="low income":
                    cl += 1
                elif j == "middle income":
                    cm += 1
                elif j == "high income":
                    ch += 1
        conn.commit()
        cur.close()
        conn.close()
        
        open_window()
        frame4=tkinter.Frame(top,bd=10,bg="papaya whip")
        frame4.pack(fill="x")
        figure1 = plt.Figure(figsize=(8,4), dpi=90)
        figure1.suptitle("Economic Division")
        ax1 = figure1.add_subplot(131)
        ax1.pie([cp,cl,cm,ch],labels = ["Below Poverty Line","Low income","Middle income","High income"],explode=[0.1,0.1,0.1,0.1],startangle=45,autopct="%1.0f%%")
        
        ax2 = figure1.add_subplot(133)
        a2 = FigureCanvasTkAgg(figure1,frame4)
        ax2.bar(1,cp,label="Below Poverty Line",color="orange")
        ax2.bar(2,cl,label="Low income")
        ax2.bar(3,cm,label="Middle income")
        ax2.bar(4,ch,label="High income")
        a2.get_tk_widget().pack()
        ax2.legend()
    
    def population_density():
        conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',database="project")
        cur=conn.cursor()
        query="select address from details"
        cur.execute(query)
        address=cur.fetchall()
        distinct_address=set(address)
        new=[x[0] for x in distinct_address]
        address_data=[]
        for i in distinct_address:
            count=0
            for j in address:
                if i==j:
                    count+=1
            address_data.append(count)
        conn.commit()
        cur.close()
        conn.close()
        
        open_window()
        frame4=tkinter.Frame(top,bd=10,bg="papaya whip")
        frame4.pack(fill="x")
        figure1 = plt.Figure(figsize=(8,4), dpi=90)
        figure1.suptitle("Area Wise Population Density")
        ax1=figure1.add_subplot(111)
        a1 = FigureCanvasTkAgg(figure1,frame4)
        colours=["red","orange","green","blue","silver","pink","grey","purple","yellow","cyan","maroon","magenta"]
        ax1.pie(address_data,labels=new,autopct='%1.0f%%',colors=colours)
        a1.get_tk_widget().pack()
    
    
        
    def name_storing():
        x=name_input.get()
        
        conn = con.connect(user="root",passwd="reckoning",host = "localhost",auth_plugin="mysql_native_password",database="project")
        cur = conn.cursor()
        query = "select * from details where Name like "+"\""+x+"\""
        cur.execute(query)
        a = cur.fetchall()
        if a==[]:
            messagebox.showerror("Name Not Found!","Sorry, Name not found in directory.")
        else:
            name_input.delete(0,tkinter.END)
            name.destroy()
            particulars(x)
        
    def age_div():
        def age_exe():
            low=age_in1.get()
            high=age_in2.get()
            if low<=high:
                age_win.destroy()
                conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',database="project")
                cur=conn.cursor()
                cur.execute("select Name,Age,Address,Phone_No from details where age>=%s and age<=%s order by Age",(low,high))
                tab=cur.fetchall()
                open_window()
                scrolling_age=scrolledtext.ScrolledText(top,width=600,height=600)
                scrolling_age.pack()
                count=len(tab)
                headers=["Name","Age","Address","Phone Number"]
                scrolling_age.insert(tkinter.INSERT,tabulate(tab,headers=headers))
                scrolling_age.insert(tkinter.INSERT,"\n_ _ _ _ _ _ _ _ _ _\n")
                scrolling_age.insert(tkinter.INSERT,("Total Count="+str(count)))
                scrolling_age.config(state=tkinter.DISABLED)
            else:
                messagebox.showerror("Oops!","Lower limit greater than upper limit")
            
        age_win=tkinter.Toplevel()
        age_win.geometry("250x250+150+150")
        label=tkinter.Label(age_win,text="Choose range of age:")
        label.pack()
        age_in1=tkinter.Spinbox(age_win,from_=1,to=150,width=15)
        age_in1.pack()
        age_in2=tkinter.Spinbox(age_win,from_=1,to=150,width=15)
        age_in2.pack()
        age_sub=tk.Button(age_win,text="Submit",command=age_exe)
        age_sub.pack()
        
        
    def result_1():
        entry_1=combo1.get()
        if entry_1=="None":
            messagebox.showinfo("None Option Selected","Please select option other than 'None' to see results.")
        elif entry_1=="See Sex Ratio":
            sex_ratio()
        elif entry_1=="Find Particulars of a person":
            global name
            name=tkinter.Toplevel()
            name.geometry("250x250+150+150")
            global name_input
            label3=tkinter.Label(name,text="Enter full name of person:")
            label3.pack()
            name_input=tkinter.Entry(name,width=15)
            name_input.focus_set()
            name_input.pack()
            name_submit=tk.Button(name,text="Submit",command=name_storing)
            name_submit.pack()
            but=tk.Button(name,text="Close",command=name.destroy)
            but.pack()
            
        elif entry_1=="See Voter List":
            voters_list()
            
        elif entry_1=="See Economic Status Graph":
            economicdatagraph()
            
        elif entry_1=="See Age Group Division":
            age_div()
            
        elif entry_1=="See Area Wise Population Density":
            population_density()
    
    window=tkinter.Tk()
    window.geometry("750x750+120+120")
    window.configure(bg="papaya whip")
    window.title("Population Survey Analysis System")
    
    labelframe = tkinter.LabelFrame(window,bg="papaya whip")
    labelframe.pack(fill="x",expand="no")
    label1 = tkinter.Label(labelframe, text="Population Survey Analysis System",font=("Arial",45,"underline"),bg="papaya whip")
    label1.pack()
    
    frame2=tkinter.Frame(window,bd=10,bg="papaya whip")
    frame2.pack(fill="x")
    
    label2=tkinter.Label(frame2,text="What would you like to do?",font=("Helvetica",20),bg="papaya whip",fg="black")
    label2.pack()
    
    combo1=tk.Combobox(frame2)
    combo1["values"]=("None","See Sex Ratio","Find Particulars of a person","See Voter List","See Economic Status Graph",
         "See Age Group Division","See Area Wise Population Density")
    combo1.current(0)
    combo1.pack()
    
    label=tkinter.Label(frame2,text="",font=("Helvetica",20),bg="papaya whip",fg="papaya whip")
    label.pack()
    
    button1=tk.Button(frame2,text="Submit",command=result_1)
    button1.pack()
    
    window.mainloop()
        
def admin():
    def admin_option():
        x=combo2.get()
        if x=="None":
            messagebox.showinfo("None Option Selected","Please select option other than 'None' to see results.")
        elif x=="Add Entry":
            input_data()
        elif x=="Delete Entry":
            del_input()
        elif x=="Change password":
            password_change()

    window1.destroy()
    window2=tkinter.Tk()
    window2.geometry("750x750+120+120")
    window2.configure(bg="papaya whip")
    window2.title("Administrator")
    
    labelframe = tkinter.LabelFrame(window2,bg="papaya whip")
    labelframe.pack(fill="x",expand="no")
    label1 = tkinter.Label(labelframe, text="Administrator",font=("Arial Bold",50,"underline"),bg="papaya whip")
    label1.pack()
    
    frame2=tkinter.Frame(window2,bd=10,bg="papaya whip")
    frame2.pack(fill="both",expand="yes")
    
    label2=tkinter.Label(frame2,text="Choose:",font=("Helvetica",30),bg="papaya whip")
    label2.pack()
    combo2=tk.Combobox(frame2)
    combo2["values"]=("None","Add Entry","Delete Entry","Change password")
    combo2.current(0)
    combo2.pack()
    
    label=tkinter.Label(frame2,text="",font=("Helvetica",20),bg="papaya whip",fg="papaya whip")
    label.pack()
    button1=tk.Button(frame2,text="Submit",command=admin_option)
    button1.pack()
    
    window2.mainloop()
    
def password_check():
    x=password_input.get()
    book=xlrd.open_workbook("NAME.xlsx")
    first_sheet=book.sheet_by_index(0)
    cell=first_sheet.cell(0,5)
    z=cell.value
    if x!=str(z):
        messagebox.showerror("Incorrect Password","Sorry, Password entered is incorrect.")
    elif x==str(z):
        password_input.delete(0,tkinter.END)
        password.destroy()
        admin()
        
def admin_click():
    global password
    password=tkinter.Toplevel()
    password.geometry("250x250+150+150")
    global password_input
    label3=tkinter.Label(password,text="Enter password:")
    label3.pack()
    password_input=tkinter.Entry(password,show="*",width=15)
    password_input.focus_set()
    password_input.pack()
    password_submit=tk.Button(password,text="Submit",command=password_check)
    password_submit.pack()
    but=tk.Button(password,text="Close",command=password.destroy)
    but.pack()


def input_data():
    def add_data():
        conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',database="project")
        cur=conn.cursor()
        query="select max(SNo) from details"
        cur.execute(query)
        a=cur.fetchone()
        new_sno=a[0]+1
        name=entry_1.get()+" "+entry_2.get()
        query="insert into details values(%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        tup=(new_sno,name,entry_3.get(),entry_4.get(),entry_5.get(),entry_6.get(),entry_7.get(),entry_8.get(),entry_9.get())
        cur.execute(query,tup)
        conn.commit()
        cur.close()
        conn.close()
        messagebox.showinfo("Successful","Data Added Successfully!")
        top.destroy()
        
    def data_check():    #Add
        p=""
        c_count=0
        a1=entry_1.get()
        if a1=="":
            p+="First Name can not be empty.\n"
        else:
            c_count+=1
        a2=entry_2.get()
        if a2=="":
            p+="Last Name can not be empty.\n"
        else:
            c_count+=1
        a4=entry_4.get()
        if len(a4)!=10:
            p+="Phone number can only be of 10 digits.\n"
        else:
            c_count+=1
        a5=entry_5.get()
        if "@" not in a5 or ".com" not in a5:
            if "@" not in a5 or ".co.in" not in a5:
                p+="Email is not proper.\n"
            else:
                c_count+=1
        else:
            c_count+=1
        a6=entry_6.get()
        if a6=="Select":
            p+="Select an address.\n"
        else:
            c_count+=1
        a7=entry_7.get()
        if a7=="Select":
            p+="Select a gender.\n"
        else:
            c_count+=1
        a8=entry_8.get()
        if a8=="Select":
            p+="Select dependency.\n"
        else:
            c_count+=1
        a9=entry_9.get()
        if a9=="Select":
            p+="Select economic status."
        else:
            c_count+=1
        if c_count==8:
            add_data()
        else:
            messagebox.showerror("Error",p)
    open_window()
    frame1=tkinter.Frame(top,bd=10,bg="papaya whip")
    frame1.pack(fill="x")
    label1=tkinter.Label(frame1,text="First Name:",bg="papaya whip")
    label1.pack()
    entry_1=tkinter.Entry(frame1,width=15)
    entry_1.pack()
    label2=tkinter.Label(frame1,text="Last Name:",bg="papaya whip")
    label2.pack()
    entry_2=tkinter.Entry(frame1,width=15)
    entry_2.pack()
    label3=tkinter.Label(frame1,text="Age:",bg="papaya whip")
    label3.pack()
    entry_3=tkinter.Spinbox(frame1,from_=1,to=150,width=15)
    entry_3.pack()
    label4=tkinter.Label(frame1,text="Phone Number:",bg="papaya whip")
    label4.pack()
    entry_4=tkinter.Entry(frame1,width=15)
    entry_4.pack()
    label5=tkinter.Label(frame1,text="Email id:",bg="papaya whip")
    label5.pack()
    entry_5=tkinter.Entry(frame1,width=15)
    entry_5.pack()
    label6=tkinter.Label(frame1,text="Address:",bg="papaya whip")
    label6.pack()
    entry_6=tk.Combobox(frame1,width=15)
    entry_6["values"]=("Select","Tulip Violet","Tata Primanti","M3M Golf Estate","Ireo Skyon","Sispal Vihar","The Legend","Emaar Mgf Palm Gardens","M3M Merlin","Ireo Victory Valley","Emaar MGF The Palm Drive","Adani M2K Oyster Grande","Devinder Vihar")
    entry_6.current(0)
    entry_6.pack()
    label7=tkinter.Label(frame1,text="Gender:",bg="papaya whip")
    label7.pack()
    entry_7=tk.Combobox(frame1,width=15)
    entry_7["values"]=("Select","Male","Female")
    entry_7.current(0)
    entry_7.pack()
    label8=tkinter.Label(frame1,text="Dependency:",bg="papaya whip")
    label8.pack()
    entry_8=tk.Combobox(frame1,width=15)
    entry_8["values"]=("Select","Dependent","Independent")
    entry_8.current(0)
    entry_8.pack()
    label9=tkinter.Label(frame1,text="Economic Status:",bg="papaya whip")
    label9.pack()
    entry_9=tk.Combobox(frame1,width=15)
    entry_9["values"]=("Select","below poverty line","low income","middle income","high income")
    entry_9.current(0)
    entry_9.pack()
    entry_submit=tk.Button(frame1,text="Submit",command=data_check)
    entry_submit.pack()        
    
def del_input():
    def del_check():
        def del_data():
            s=del_entry2.get()
            if s=="":
                messagebox.showerror("Error","Nothing entered!")
            else:
                conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',database="project")
                cur=conn.cursor()
                query="select max(SNo) from details"
                cur.execute(query)
                m=cur.fetchone()
                if s[0].isalpha() or int(s)>m[0] or int(s)<=0:
                    messagebox.showerror("Error","Invalid SNo")
                else:
                    query="delete from details where SNo='%s'"%(s,)
                    cur.execute(query)
                    query="update details set sno=sno-1 where sno>%s"%(s,)
                    cur.execute(query)
                    conn.commit()
                    cur.close()
                    conn.close()
                    messagebox.showinfo("Done","Entry removed!")
                    top.destroy()
        
        w=del_entry1.get()
        conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',database="project")
        cur=conn.cursor()
        query="select * from details where name like '%s'"%(w,)
        cur.execute(query)
        x=("SNo:%s , Name:%s , Age:%s , Phone Number:%s , Email:%s")
        a=cur.fetchall()
        if a==[]:
            messagebox.showerror("Name Not Found","Name not found in Directory!")
        else:
            for i in a:
                label=tkinter.Label(frame1,text=x%(i[0],i[1],i[2],i[3],i[4]),bg="papaya whip")
                label.pack()
            label2=tkinter.Label(frame1,text="Enter SNo of entry you want to remove:",bg="papaya whip")
            label2.pack()
            del_entry2=tkinter.Entry(frame1,width=15)
            del_entry2.pack()
            del_submit2=tk.Button(frame1,text="Submit",command=del_data)
            del_submit2.pack()
        conn.commit()
        cur.close()
        conn.close()
    
    open_window()
    frame1=tkinter.Frame(top,bd=10,bg="papaya whip")
    frame1.pack(fill="x")
    label1=tkinter.Label(frame1,text="Full Name:",bg="papaya whip")
    label1.pack()
    del_entry1=tkinter.Entry(frame1,width=15)
    del_entry1.pack()
    del_submit=tk.Button(frame1,text="Submit",command=del_check)
    del_submit.pack()


def password_change():
    def change_check():
        x=entry_1.get()
        book=xlrd.open_workbook("NAME.xlsx")
        first_sheet=book.sheet_by_index(0)
        cell=first_sheet.cell(0,5)
        if str(cell.value)==x:
            y=entry_2.get()
            if len(y)<8:
                messagebox.showerror("Error","Password length should be more than 8!")
            else:
                if y==entry_3.get():
                    count_alpha=count_num=0
                    for i in y:
                        if i.isalpha():
                            count_alpha+=1
                        elif i.isdigit():
                            count_num+=1
                    
                    if count_alpha<1:
                        messagebox.showerror("Error","Password should have at atleast 1 alphabet")
                    elif count_num<1:
                        messagebox.showerror("Error","Password should have at atleast 1 number")
                    else:
                        book=xlrd.open_workbook("NAME.xlsx")
                        first_sheet=book.sheet_by_index(0)
                        wbook=copy(book)
                        w_sheet = wbook.get_sheet(0)
                        w_sheet.write(0,5,y)
                        wbook.save("NAME.xlsx") 
                        messagebox.showinfo("Success","Password Changed Successfully")
                        top.destroy()
                else:
                    messagebox.showerror("Error","New password and Re-enter New Password don't match")
        else:
            messagebox.showerror("Error","Incorrect Current Password")
    open_window()
    frame1=tkinter.Frame(top,bd=10,bg="papaya whip")
    frame1.pack(fill="x")
    label=tkinter.Label(frame1,text="Password must be of 8 characters or more\nMust contain atleast 1 Number and 1 alphabet",bg="papaya whip")
    label.pack()
    label1=tkinter.Label(frame1,text="Current Password:",bg="papaya whip")
    label1.pack()
    entry_1=tkinter.Entry(frame1,show="*",width=15)
    entry_1.pack()
    label2=tkinter.Label(frame1,text="New Password",bg="papaya whip")
    label2.pack()
    entry_2=tkinter.Entry(frame1,show="*",width=15)
    entry_2.pack()
    label3=tkinter.Label(frame1,text="Re-enter New Password:",bg="papaya whip")
    label3.pack()
    entry_3=tkinter.Entry(frame1,show="*",width=15)
    entry_3.pack()
    pass_submit=tk.Button(frame1,text="Submit",command=change_check)
    pass_submit.pack()
    

window1=tkinter.Tk()
window1.geometry("750x750+120+120")
window1.configure(bg="pale turquoise")
window1.title("Selection Screen")

labelframe = tkinter.LabelFrame(window1,bg="pale turquoise")
labelframe.pack(fill="x",expand="no")
label1 = tkinter.Label(labelframe, text="Population Survey Analysis System",font=("Arial",45,"underline"),bg="pale turquoise")
label1.pack()

frame2=tkinter.Frame(window1,bd=10,bg="pale turquoise")
frame2.pack(fill="both",expand="yes")

label2=tkinter.Label(frame2,text="Choose:",font=("Helvetica",30),bg="pale turquoise")
label2.pack()

button1=tk.Button(frame2,text="User",command=user)
button1.pack(ipady=125, ipadx=125)
button2=tk.Button(frame2,text="Admin",command=admin_click)
button2.pack(ipady=125, ipadx=125, side="bottom")

window1.mainloop()