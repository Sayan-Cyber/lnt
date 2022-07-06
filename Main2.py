#NECESSARY MODULES____________________________________________________________________________________________________________________________

'''These are the necessary modules needed to properly run this. Here Tkinter is not imported 
as '*' because that is a bad practice. Here only necessary modules from tkinter has been imported '''

"""PEP 8 is not followed strictly due to security reasons"""


import math 
import os
import sys
from tkinter.ttk import Combobox
import webbrowser
from tkinter import (BOTH, CENTER, LEFT, RAISED, RIDGE, RIGHT, SUNKEN, FLAT,
                    TOP, TRUE, Button, Canvas, Frame,
                    Label, N, PhotoImage, StringVar, Tk, Toplevel, W,
                    font, messagebox,Entry)
from tkinter.filedialog import asksaveasfilename
import pyglet

from PIL import Image, ImageTk

'''Necessary fonts has been added using pyglet module'''
pyglet.font.add_file('Additional_File\JosefinSans-Bold.ttf')
pyglet.font.add_file('Additional_File\ReadexPro-Medium.ttf')

counter=0
w='#ffffff'
b='#000000'
lbf='Readex Pro Medium',12


import openpyxl

from win32com import client


print


xl = openpyxl.load_workbook("test1.xlsx")

def calculations ():                                                                  #function for the save button in the main sheet___________________________________________________________
    wname = wtp_name.get()     
    dname = doc_name.get()
    sname = c_name.get()
    tname = t_name.get()
    date= dat_e.get()
    dby= p_by.get()
    cby= c_by.get()
    doca=docu_name_a.get()
    docb=docu_name_b.get()
    docc=docu_name_c.get() 
    docd=docu_name_d.get()
    doce=docu_name_e.get()
    docf=docu_name_f.get()
    docg=docu_name_g.get()
    doch=docu_name_h.get()

    os.chdir(sys.path[0])
 
    sheet = xl['EP-NAVADA PS']
    sheet['D5'] =wname
    sheet['D7'] = tname
    sheet['Z8'] =cby
    sheet['V8'] = dby
    sheet['AF6'] = date
    sheet['AH8'] = sname

      
    amp = float(i.get())
    teev = float(tee.get())
    neev  = nee.get() 
    legsv= legs.get()
    asrv= asr.get()
    mat= matr.get()

    doclist=[dname,doca,docb,docc,docd ]
    docstring= "-".join(doclist)
    docstring2=[docstring,doce,docf,docg,doch]
    docstringf="".join(docstring2)
    if mat=="AL":
        k= 126
        b=228 
        delta=.0025
        q=0.000138
    elif mat=="CU":
        k=205
        b=234
        delta=.00345
        q=0.000138
    elif mat=="GI":
        k= 79.14  
        b=202
        delta=.0038
        q=0.000138
   
    s= 1000*((amp*1)/k)
      
    teevf= (s+(s*teev/100))
    os.chdir(sys.path[0])
   
    sheet['V22'] = amp
    sheet['V12'] = mat
    sheet['V19'] = neev
    sheet['V20'] = legsv
    sheet['V6'] = docstringf
    
    
    sheet['V25'] = b
    sheet['V26'] = q
    sheet['V24'] = delta

    

    sheet['V39'] = format(k.real,".2f")
    sheet['V40'] = format(s.real,".2f")
    
    sheet['V41'] =teev
    sheet['V42'] =teevf
    
    if teevf<= 75:
        strip="25x3"
    elif teevf<= 300 and teevf>75 :
        strip="50x6"
    elif teevf<= 500 and teevf>300:
        strip="50x10"
    elif teevf<= 650 and teevf>500:
        strip="65x10"
    elif teevf<= 750 and teevf>650:
        strip="75x10"
        
    else:
        messagebox.showerrror("Error","Please Check Your Calculations Again")
        
    sheet['V45'] = strip
    
    sheet.title="EP-xxxx-PS"
    
def savexl():
    calculations()
    res =messagebox.askquestion("Save File", "Are you Sure you want to save this file?")
    if res == 'yes' :
        files = [('Excel Document(.xlsx)','*.xlsx')]
        file = asksaveasfilename(filetypes = files, defaultextension=files)
        if file=='':
            messagebox.showinfo('Error', 'Your File was not Saved :(')
        else:
            
            xl.save(file)
            
            messagebox.showinfo("Save","Your Calculations was exported successfully!")
            
    else:
        messagebox.showinfo('Error','Your Document was not saved')   
    xl.close()


def pdf():
    try:
        # Open Microsoft Excel
        calculations()
        xl.save("temp.xlsx")
        excel = client.Dispatch("Excel.Application")
        # Read Excel File"C:\Users\20323801\Documents\OneNote Notebooks"
        sheets = excel.Workbooks.Open("temp")
        work_sheets = sheets.Worksheets[5]
        res =messagebox.askquestion("Save File", "Are you Sure you want to save this file?")
        if res == 'yes' :
            files = [('PDF Document(.pdf)','*.pdf')]
            file = asksaveasfilename(filetypes = files, defaultextension=files)
            if file=='':
                
                messagebox.showinfo('Error', 'Your File wad not Saved :(')
            else:
                work_sheets.ExportAsFixedFormat(0, file)
                messagebox.showinfo("Save","Your Calculations was exported successfully!")
        else:
            messagebox.showinfo('Error','Your Document was not saved')
    except:
        messagebox.showinfo('Error','This Option is not available Yet! ')
    xl.close()
   
                                                                          

root= Tk()
root.title( 'Earthing Calculation')
root.state('zoomed')
root.geometry('1400x800+100+100')
root.minsize(1260,720)
root.configure(bg=w)
bgimg= PhotoImage(file='Additional_File\\bg.png')

dash= PhotoImage(file='Additional_File\\icons\dashb.png')
photo3= dash.subsample(1,1)

canvas1 = Canvas( root, width = 1920,height = 1080,bd=0,highlightthickness=0, background="white")
  
canvas1.pack(fill=BOTH,expand=TRUE)
  
canvas1.create_image( 0, 0, image =bgimg,anchor = "nw")

# _____________PARAMETER SUBMENUS_________________________________________________________

def clicked():
    root.counter += 1
    if root.counter == 3:
         webbrowser.open(url='https://bit.ly/3dAlL5T')
         root.counter=0
    
  
Label(root,text="Earthing & Lightning Calculations",width=200,bg='#6F00C7',font=('Readex Pro Medium',18),borderwidth=0,compound=LEFT).place(relx=.5,y=25,anchor=CENTER)

Button(root,text='Dashboard',image=photo3,compound=LEFT,width=196,height=45,bg='#6F00C7',borderwidth=0,font=('Readex Pro Medium',18),command=clicked,activebackground='#23D2FF').place(x=0,y=0)
  
TEMP= Frame(root)


frame2= Frame(root, bg="WHITE",borderwidth=1,relief=FLAT,height=300,padx=40,pady=35,highlightbackground="blACK", highlightthickness=2)

frame2.place(relx=0.5, rely=0.3, anchor=CENTER)


frame4= Frame(root, bg="WHITE",borderwidth=0,relief=FLAT,height=300,padx=40,pady=50)

frame4.place(relx=0.5, rely=0.75, anchor=CENTER)



msheet= PhotoImage(file='Additional_File\icons\sheet.png')
photo15= msheet.subsample(2,2)
instrume= PhotoImage(file='Additional_File\icons\para.png')
photo1= instrume.subsample(1,1)
write= PhotoImage(file="Additional_File\\icons\\name.png")
photo9= write.subsample(2,2)
doc= PhotoImage(file='Additional_File\icons\doc.png')
photo8= doc.subsample(2,2)
c= PhotoImage(file='Additional_File\icons\client.png')
photo7= c.subsample(2,2)
rev= PhotoImage(file='Additional_File\\icons\\rev.png')
photo6= rev.subsample(2,2)
date= PhotoImage(file='Additional_File\icons\date.png')
photo5= date.subsample(2,2)
pre= PhotoImage(file='Additional_File\icons\pre.png')
photo4= pre.subsample(2,2)
check= PhotoImage(file='Additional_File\icons\check.png')
photo10= check.subsample(2,2)
app= PhotoImage(file='Additional_File\\icons\\app.png')
photo11= app.subsample(2,2)
state= PhotoImage(file='Additional_File\icons\state.png')
photo12= state.subsample(2,2)
place= PhotoImage(file='Additional_File\icons\place.png')
photo13= place.subsample(2,2)
star= PhotoImage(file='Additional_File\icons\star.png')
photo99= star.subsample(2,2)


idateicon= Image.open('Additional_File\\icons\idate.png')
coverphoto4 = ImageTk.PhotoImage(idateicon)

Label(frame2,text='General', font=('Readex Pro Medium',20), fg='#247881' ,bg=w, image=photo15,compound=LEFT).grid(row=0,column=0,sticky=W)
lb2=Label(frame2,text='Project Name :', font=('Readex Pro Medium',14), fg='#6F00C7' ,bg=w, image=photo9,compound=LEFT)
lb2.grid(row=1,column=0,padx=10, sticky=W)
lb4=Label(frame2,text='Title :', font=('Readex Pro Medium',14), fg=	'#6F00C7',bg=w , image=photo7,compound=LEFT)
lb4.grid(row=2,column=0,pady=5,padx=10, sticky=W)
lb3=Label(frame2,text='Sheet :', font=('Readex Pro Medium',14), fg='#6F00C7',bg=w , image=photo10,compound=LEFT)
lb3.grid(row=3,column=0,pady=5,padx=10, sticky=W)
lb6=Label(frame2,text='Date :', font=('Readex Pro Medium',14), fg='#6F00C7',bg=w , image=photo5,compound=LEFT)
lb6.grid(row=1,column=2,pady=5,padx=25, sticky=W)
lb7=Label(frame2,text='Designed By :', font=('Readex Pro Medium',14), fg='#6F00C7',bg=w , image=photo4,compound=LEFT)
lb7.grid(row=2,column=2,pady=5,padx=25, sticky=W)
lb8=Label(frame2,text='Checked By :', font=('Readex Pro Medium',14), fg='#6F00C7',bg=w , image=photo10,compound=LEFT)
lb8.grid(row=3,column=2,pady=5,padx=25, sticky=W)



wtp_name = Entry(frame2,width=20,borderwidth=.5, relief=FLAT,highlightcolor='#6F00C7',highlightthickness=1)
wtp_name.grid(row=1,column=1,pady=5,padx=1)
t_name = Entry(frame2,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#6F00C7',highlightthickness=1)
t_name.grid(row=2,column=1, pady=5,padx=1)
c_name = Entry(frame2,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#6F00C7',highlightthickness=1)
c_name.grid(row=3,column=1, pady=5,padx=1)
dat_e= Entry(frame2,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#6F00C7',highlightthickness=1)
dat_e.grid(row=1,column=3, pady=5,padx=1)
p_by= Entry(frame2,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#6F00C7',highlightthickness=1)
p_by.grid(row=2,column=3, pady=5,padx=1)
c_by= Entry(frame2,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#6F00C7',highlightthickness=1)
c_by.grid(row=3,column=3, pady=5,padx=1)

doc_name = Entry(frame2,width=15,borderwidth=.5, relief=FLAT,highlightcolor='#6F00C7',highlightthickness=1)
doc_name.grid(row=4, column=1,pady=5,padx=1,sticky=W)




 
  
  
  
  
  
  
  
  
  
Label(frame4,text='Inputs', font=('Readex Pro Medium',20), fg='#247881' ,bg=w, image=photo15,compound=LEFT).grid(row=0,column=0,sticky=W)
lb2=Label(frame4,text='Max. Fault Current:', font=('Readex Pro Medium',14), fg='#6F00C7' ,bg=w, image=photo99,compound=LEFT)
lb2.grid(row=1,column=0,padx=10, sticky=W)
lb4=Label(frame4,text='Earth Strip Material :', font=('Readex Pro Medium',14), fg=	'#6F00C7',bg=w , image=photo99,compound=LEFT)
lb4.grid(row=2,column=0,pady=5,padx=10, sticky=W)
lb5=Label(frame4,text='Allowances in Cross Sectional Area:', font=('Readex Pro Medium',14), fg=	'#6F00C7',bg=w , image=photo99,compound=LEFT)
lb5.grid(row=3,column=0,pady=5,padx=10, sticky=W)
lb6=Label(frame4,text='No. of Earth Electrodes:', font=('Readex Pro Medium',14), fg='#6F00C7',bg=w , image=photo99,compound=LEFT)
lb6.grid(row=1,column=2,pady=5,padx=25, sticky=W)
lb7=Label(frame4,text='Length of Earth Grid strip:', font=('Readex Pro Medium',14), fg='#6F00C7',bg=w , image=photo99,compound=LEFT)
lb7.grid(row=2,column=2,pady=5,padx=25, sticky=W)
lb8=Label(frame4,text='Average Soil resistivity :', font=('Readex Pro Medium',14), fg='#6F00C7',bg=w , image=photo99,compound=LEFT)
lb8.grid(row=3,column=2,pady=5,padx=25, sticky=W)





      

      
i = Entry(frame4,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#00FFC6',highlightthickness=1)
i.grid(row=1,column=1,pady=5,padx=1)
mat= StringVar()
matr= Combobox(frame4,textvariable=mat,width=20)
matr.grid(row=2,column=1, pady=5,padx=1)
tee= Entry(frame4,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#00FFC6',highlightthickness=1)
tee.grid(row=3,column=1, pady=5,padx=1)
nee= Entry(frame4,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#00FFC6',highlightthickness=1)
nee.grid(row=1,column=3, pady=5,padx=1)
legs= Entry(frame4,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#00FFC6',highlightthickness=1)
legs.grid(row=2,column=3, pady=5,padx=1)
asr= Entry(frame4,width=20,borderwidth=.5, relief=SUNKEN,highlightcolor='#00FFC6',highlightthickness=1)
asr.grid(row=3,column=3, pady=5,padx=1)




  
  
  
    

dnumicon= PhotoImage(file='Additional_File\\icons\dnum.png')
coverphoto5 = dnumicon.subsample(2,2)
l5=Label(frame2, text='Document No :',font=('Readex Pro Medium',14), fg='#6F00C7' ,bg=w,image=coverphoto5,compound=LEFT)
    
l5.image= coverphoto5
l5.grid(row=4,column=0,sticky=W,padx=10)


frame1= Frame(frame2, bg=w)
frame1.place(x=270,y=197)



docu_name_a= StringVar()
docuname_a= Combobox(frame1,width=4, textvariable=docu_name_a)
docuname_a['values']= ('M','E','C','I','P')
docuname_a.grid(row=1,column=1,pady=5,sticky=W, padx=0)
docu_name_b= StringVar()
docuname_b= Combobox(frame1,width=4, textvariable=docu_name_b)
docuname_b['values']= ('W S',)
docuname_b.grid(row=1,column=3,pady=5,sticky=W, padx=0)
    
docu_name_c= StringVar()
docuname_c= Combobox(frame1,width=4, textvariable=docu_name_c)
docuname_c['values']= ('C W'," W T")
docuname_c.grid(row=1,column=5,pady=5,sticky=W, padx=0)
    
docu_name_d= StringVar()
docuname_d= Combobox(frame1,width=4, textvariable=docu_name_d)
docuname_d['values']= ('D C')
docuname_d.grid(row=1,column=7,pady=5,padx=0, sticky=W)
    
docu_name_e= StringVar()
docuname_e= Combobox(frame1,width=1, textvariable=docu_name_e)
docuname_e['values']= ('0','1','2','3','4','5','6','7','8','9')
docuname_e.grid(row=1,column=9,pady=5,padx=1)
docu_name_f= StringVar()
docuname_f= Combobox(frame1,width=1, textvariable=docu_name_f)
docuname_f['values']= ('0','1','2','3','4','5','6','7','8','9')
docuname_f.grid(row=1,column=10,pady=5,padx=1)
docu_name_g= StringVar()
docuname_g= Combobox(frame1,width=1, textvariable=docu_name_g)
docuname_g['values']= ('0','1','2','3','4','5','6','7','8','9')
docuname_g.grid(row=1,column=11,pady=5,padx=1)
docu_name_h= StringVar()
docuname_h= Combobox(frame1,width=1, textvariable=docu_name_h)
docuname_h['values']= ('0','1','2','3','4','5','6','7','8','9')
docuname_h.grid(row=1,column=12,pady=5,padx=1)

Label(frame1,text="-", font=('Readex Pro Medium',18), fg='#0E185F' ,bg=w).grid(row=1,column=0)
Label(frame1,text="-", font=('Readex Pro Medium',18), fg='#0E185F' ,bg=w).grid(row=1,column=2)
Label(frame1,text="-", font=('Readex Pro Medium',18), fg='#0E185F' ,bg=w).grid(row=1,column=4)
Label(frame1,text="-", font=('Readex Pro Medium',18), fg='#0E185F' ,bg=w).grid(row=1,column=6)
Label(frame1,text="-", font=('Readex Pro Medium',18), fg='#0E185F' ,bg=w).grid(row=1,column=8)
    
exit= PhotoImage(file='Additional_File\icons\exit.png')
photoimage1= exit.subsample(2,2)
savee= PhotoImage(file='Additional_File\icons\save.png')
photoimage2= savee.subsample(2,2)

Button(root,fg="#000000",text='Exit',font=('Josefin Sans',10),image=photoimage1,compound=RIGHT,bg='#ffffff', command=root.destroy,borderwidth=0,cursor='hand2').place(x=1850,rely=.1,anchor=CENTER)


Button(frame4,text="Export Excel",fg="#000000",font=('Josefin Sans',10, font.BOLD),image=photoimage2,compound=RIGHT,cursor='hand2',bg=w, command=savexl,borderwidth=0).place(relx=0.2,rely=1.01,anchor=CENTER)
Button(frame4,text="Export PDF",fg="#000000",font=('Josefin Sans',10, font.BOLD),image=photoimage2,compound=RIGHT,cursor='hand2',bg="#6F00C7", command=pdf,borderwidth=0).place(relx=0.6,rely=1.01,anchor=CENTER)

image1= PhotoImage(file='Additional_File\We.png')
image1.subsample(1,1)

image301 = Image.open("Additional_File\icons\\boq.png")
test301 = ImageTk.PhotoImage(image301)
Label(root,image=image1,bg='#6F00C7').place(relx=.86,y=2)
copyrightLabel= canvas1.create_text(960,990,text='© L&T Construction (WET IC-WSD) 2022',font=('Tahoma',9,font.BOLD),fill="#000000")
root.mainloop()       











