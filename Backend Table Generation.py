#2
import mysql.connector as con
import xlrd
import random

conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password')
cur=conn.cursor()
query="create database if not exists Project"
cur.execute(query)
conn.commit()
cur.close()
conn.close()

conn=con.connect(user="root",password="reckoning",host="localhost",auth_plugin='mysql_native_password',db="Project")
cur=conn.cursor()
query="create table if not exists Details(SNo int,Name varchar(30),Age int,Phone_No varchar(50),Email varchar(100),Address varchar(100),Gender Varchar(10),Dependency varchar(20),Economic_Status varchar(30))"
cur.execute(query)

l=[]#First name
m=[]#Surname
n=[]#Fullname
o=[]#Age
p=[]#Dependency
q=[]#Gender
r=[]#Number
s=[]#Address
t=[]#Email
u=[]#Serial number
status=[]#status
book=xlrd.open_workbook("NAME.xlsx")
first_sheet=book.sheet_by_index(0)
for i in range(200):
    a=random.randint(0,429)
    cell=first_sheet.cell(a,0)
    l.append(cell.value.strip())
    cell2=first_sheet.cell(a,1)
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
    k=''
    A=9
    for j in range(10):
        k+=str(A)
        A=random.randint(0,9)
    r.append(int(k))
    B=l[i].lower()+str(o[i])+"@gmail.com"
    t.append(B)
    u.append(i+1)
    pov = ["below poverty line","middle income","low income","high income"]
    ran = random.randint(0,3)
    dep = pov[ran]
    status.append(dep)
    
for i in o:
    if i<=20:
        a="Dependent"
    elif i>=65:
        a="Dependent"
    else:
        a="Independent"
    p.append(a)

for i in range(200):
    query=("insert into Details values(%s,%s,%s,%s,%s,%s,%s,%s,%s)")
    tup=(u[i],n[i],o[i],r[i],t[i],s[i],q[i],p[i],status[i])
    cur.execute(query,tup)
conn.commit()
cur.close()
conn.close()
