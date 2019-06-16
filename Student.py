import xlrd
from Tkinter import *
var=""

locat="data.xlsx"
workbook=xlrd.open_workbook(locat)
sheet=workbook.sheet_by_index(0)

def dispdet(x):
    x=int(x)
    deta=Tk()
    deta.title("Student details")
    text=Text(deta,font="Calibri")
    for i in range(0,sheet.ncols):
        text.insert(INSERT,sheet.cell_value(0,i))
        text.insert(INSERT,": ")
        text.insert(INSERT,sheet.cell_value(x,i))
        text.insert(INSERT,"\n")
    text.config(state="disabled")
    text.pack()
def disprn(rno):
    rno=int(rno)
    for i in range (0,sheet.nrows):
        if(rno==sheet.cell_value(i,0)):
           dispdet(i)
           break
    else:
        deta=Tk()
        deta.title("Student details")
        text=Text(deta,font="Calibri")
        text.insert(INSERT,"Student record does not exist")
        text.config(state="disabled")
        text.pack()
def submark(x):
    deta=Tk()
    deta.title("Subject Marklist")
    text=Text(deta,font="Calibri")
    d=dict()
    l=["O[S]","A+","A","B+","B","C","F","FE"]
    for i in range(0,sheet.ncols):
        if(x==sheet.cell_value(0,i)):
           n=i
           break
    for i in range(1,sheet.nrows):
        d[sheet.cell_value(i,9)]=sheet.cell_value(i,n)
    text.insert(INSERT,"The mark list for the subject ")
    text.insert(INSERT,x)
    text.insert(INSERT,":\n")
    for i in l:
        for k,v in d.iteritems():
            if (v==i):
                text.insert(INSERT,"Name:")
                text.insert(INSERT,k)
                text.insert(INSERT,"\n")
                text.insert(INSERT,"Grade:")
                text.insert(INSERT,v)
                text.insert(INSERT,"\n")
    text.config(state="disabled")
    text.pack()
def subtop(x):
    deta=Tk()
    deta.title("Subject-wise analysis")
    text=Text(deta,font="Calibri")
    d=dict()
    for i in range(0,sheet.ncols):
        if(x==sheet.cell_value(0,i)):
           n=i
           break
    for i in range(1,sheet.nrows):
        d[sheet.cell_value(i,0)]=sheet.cell_value(i,n)
    text.insert(INSERT,"The toppers of the subject are \n")
    for k,v in d.iteritems():
        if(v=="O[S]"):
            text.insert(INSERT,"\nRoll No:")
            text.insert(INSERT,k)
            text.insert(INSERT,"\nReg No:")
            text.insert(INSERT,sheet.cell_value(int(k),1))
            text.insert(INSERT,"\nName:")
            text.insert(INSERT,sheet.cell_value(int(k),9))
    text.insert(INSERT,"\n\nThe students who obtained average marks are \n")
    for k,v in d.iteritems():
        if(v=="B+"):
            text.insert(INSERT,"\nRoll No:")
            text.insert(INSERT,k)
            text.insert(INSERT,"\nReg No:")
            text.insert(INSERT,sheet.cell_value(int(k),1))
            text.insert(INSERT,"\nName:")
            text.insert(INSERT,sheet.cell_value(int(k),9))
    text.insert(INSERT,"\n\nThe students who obtained least marks are \n")
    for k,v in d.iteritems():
        if(v=="F"):
            text.insert(INSERT,"\nRoll No:")
            text.insert(INSERT,k)
            text.insert(INSERT,"\nReg No:")
            text.insert(INSERT,sheet.cell_value(int(k),1))
            text.insert(INSERT,"\nName:")
            text.insert(INSERT,sheet.cell_value(int(k),9))
    text.config(state="disabled")
    text.pack()

def subpass(x):
    deta=Tk()
    deta.title("Pass-Fail Percentage")
    text=Text(deta,font="Calibri")
    l=list()
    count=0
    for i in range(0,sheet.ncols):
        if(x==sheet.cell_value(0,i)):
           n=i
           break
    for i in range(1,sheet.nrows):
        if(sheet.cell_value(i,n)=="FE"):
            count=count+1
    pcount=63-count
    ppg=float((pcount/63.0))*100
    fpg=float((count/63.0))*100
    text.insert(INSERT,"The pass percentage is ")
    text.insert(INSERT,ppg)
    text.insert(INSERT,"\nThe fail percentage is ")
    text.insert(INSERT,fpg)
    text.config(state="disabled")
    text.pack()
    

def ddf():
    df=Tk()
    df.title("Student")
    df.geometry("300x200")
    text=Text(df,font="Calibri",height="1")
    text.insert(INSERT,"   Enter the roll no of the student")
    text.config(state="disabled")
    text.pack(pady=20)
    E1 = Entry(df)
    E1.pack(pady=20)
    s=Button(df,text="Submit",command=lambda:disprn(E1.get()),textvariable=StringVar)
    s.pack(pady=20)
    


def ml():
    mll=Tk()
    mll.title("Subject")
    mll.geometry("800x700")
    text=Text(mll,font="Calibri",height="1")
    text.insert(INSERT,"\t\t\t\tSelect the subject")
    text.config(state="disabled")
    text.pack(pady=20)
    s1=Button(mll,text="S1MA101",command=lambda:submark("S1MA101"))
    s2=Button(mll,text="S1CY100",command=lambda:submark("S1CY100"))
    s3=Button(mll,text="S1BE100",command=lambda:submark("S1BE100"))
    s4=Button(mll,text="S1BE10105",command=lambda:submark("S1BE10105"))
    s5=Button(mll,text="S1BE103",command=lambda:submark("S1BE103"))
    s6=Button(mll,text="S1EC100",command=lambda:submark("S1EC100"))
    s7=Button(mll,text="S1CY110",command=lambda:submark("S1CY110"))
    s8=Button(mll,text="S1CS110",command=lambda:submark("S1CS110"))
    s9=Button(mll,text="S1EC110",command=lambda:submark("S1EC110"))
    s1.pack(pady=20)
    s2.pack(pady=20)
    s3.pack(pady=20)
    s4.pack(pady=20)
    s5.pack(pady=20)
    s6.pack(pady=20)
    s7.pack(pady=20)
    s8.pack(pady=20)
    s9.pack(pady=20)
 
def tl():
    tll=Tk()
    tll.title("Subject")
    tll.geometry("800x700")
    text=Text(tll,font="Calibri",height="1")
    text.insert(INSERT,"\t\t\t\tSelect the subject")
    text.config(state="disabled")
    text.pack(pady=20)
    s1=Button(tll,text="S1MA101",command=lambda:subtop("S1MA101"))
    s2=Button(tll,text="S1CY100",command=lambda:subtop("S1CY100"))
    s3=Button(tll,text="S1BE100",command=lambda:subtop("S1BE100"))
    s4=Button(tll,text="S1BE10105",command=lambda:subtop("S1BE10105"))
    s5=Button(tll,text="S1BE103",command=lambda:subtop("S1BE103"))
    s6=Button(tll,text="S1EC100",command=lambda:subtop("S1EC100"))
    s7=Button(tll,text="S1CY110",command=lambda:subtop("S1CY110"))
    s8=Button(tll,text="S1CS110",command=lambda:subtop("S1CS110"))
    s9=Button(tll,text="S1EC110",command=lambda:subtop("S1EC110"))
    s1.pack(pady=20)
    s2.pack(pady=20)
    s3.pack(pady=20)
    s4.pack(pady=20)
    s5.pack(pady=20)
    s6.pack(pady=20)
    s7.pack(pady=20)
    s8.pack(pady=20)
    s9.pack(pady=20)   

def pf():
    pff=Tk()
    pff.title("Subject")
    pff.geometry("800x700")
    text=Text(pff,font="Calibri",height="1")
    text.insert(INSERT,"\t\t\t\tSelect the subject")
    text.config(state="disabled")
    text.pack(pady=20)
    s1=Button(pff,text="S1MA101",command=lambda:subpass("S1MA101"))
    s2=Button(pff,text="S1CY100",command=lambda:subpass("S1CY100"))
    s3=Button(pff,text="S1BE100",command=lambda:subpass("S1BE100"))
    s4=Button(pff,text="S1BE10105",command=lambda:subpass("S1BE10105"))
    s5=Button(pff,text="S1BE103",command=lambda:subpass("S1BE103"))
    s6=Button(pff,text="S1EC100",command=lambda:subpass("S1EC100"))
    s7=Button(pff,text="S1CY110",command=lambda:subpass("S1CY110"))
    s8=Button(pff,text="S1CS110",command=lambda:subpass("S1CS110"))
    s9=Button(pff,text="S1EC110",command=lambda:subpass("S1EC110"))
    s1.pack(pady=20)
    s2.pack(pady=20)
    s3.pack(pady=20)
    s4.pack(pady=20)
    s5.pack(pady=20)
    s6.pack(pady=20)
    s7.pack(pady=20)
    s8.pack(pady=20)
    s9.pack(pady=20)   

    
main=Tk()
main.title("Datacheq")
main.geometry('1920x1080')
text=Text(main,font=("Cambria",80,"bold"),fg="#FFFFFF",bg="#67C8FF",height="1",padx="475",pady="100")
text.insert(INSERT,"DATACHEQ")
text.config(state="disabled")
dd=Button(main,text="Details of a particular student",font=("Calibri",25),command=ddf,bg="#67C8FF")
mt=Button(main,text="Mark list of a particular subject in ranking order",font=("Calibri",25),command=ml,bg="#67C8FF")
tal=Button(main,text="Performance sheet of students in a subject",font=("Calibri",25),command=tl,bg="#67C8FF")
pf=Button(main,text="Pass percentage and fail percentage of a particular subject",font=("Calibri",25),command=pf,bg="#67C8FF")
label = Label( main, textvariable=var, relief=RAISED )

text.pack()
dd.pack(pady="20")
mt.pack(pady="20")
tal.pack(pady="20")
pf.pack(pady="20")
main.mainloop()
