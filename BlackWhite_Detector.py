#! /usr/bin/python3
from tkinter import*
import tkinter as tk
from  tkinter import ttk, messagebox
import cv2
from PIL import Image,ImageTk
import numpy as np
import RPi.GPIO as GPIO
import time
import csv
from datetime import datetime
import os
import pandas as pd

###############Writetocsv########################

def writetocsv(transaction, filename ='transaction.csv'):
    with open(filename, 'a', newline ='', encoding='utf-8') as file:
        fw = csv.writer(file)
        fw.writerow(transaction)

#############GUI part###################
GUI = tk.Tk()

W = 800
H = 420

#เปิดโปรแกรมทุกครั้งจะอยู่ตรงกลางหน้าจอ
#หาค่าขนาดหน้าจอที่ตั้งค่า Screen resolution
MW =GUI.winfo_screenwidth()
MH = GUI.winfo_screenheight()
SX = (MW/2) - (W/2) #Start X
SY = (MH/2) - (H/2)#Start Y
SY = SY #diff up



GUI.geometry('{}x{}+{:.0f}+{:.0f}'.format(W,H,SX,SY))
GUI.title('Black&White Detector')

Font1 = ('consolas',10,'bold')
Font2 = ('consolas',8,'bold')
Font3 = ('consolas',6,'bold')

canvas1=tk.Canvas(GUI,width=300, height=370, bg='light blue')
canvas1.place(x=0,y=80)

canvas2=tk.Canvas(GUI,width=300, height=420)
canvas2.place(x=300,y=80)


#Item name frame
frame_1 = tk.Frame(GUI, width=200, height=100, bg='purple')
frame_1.place(x=600, y=0)

###thresh bar frame
frame_2 = tk.Frame(GUI, width=300, height=80, bg='pink')
frame_2.place(x=0, y=0)

###Set up ROI bar frame
frame_3 = tk.Frame(GUI, width=200, height=350, bg='orange')
frame_3.place(x=600, y=80)

###OK/NG text frame
frame_4 = tk.Frame(GUI, width=100, height=30, bg='blue')
frame_4.place(x=400, y=20)

#Transaction ID

v_transaction = StringVar()
trstamp = datetime.now().strftime('%Y%m%d%H%M%S') #generate trannsaction
v_transaction.set(trstamp)

#Label Transaction
LTR = Label(GUI, textvariable=v_transaction).place()
# defect1=0
# defect2 =0

def AddTransaction():
    global A, transaction
    #writetocsv('transaction.csv')
    stamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    transaction = v_transaction.get()
    itemname = v_itemname.get()
    v_defect1 = 'NG'
    
    A = [stamp, itemname, v_defect1]

    if v_rbutton.get()==1:
        writetocsv(A,'transaction.csv')

#####History window####################

def HistoryWindow(event):

    HIS =Toplevel() # คำสั่งคล้ายๆกับ GUI =Tk()
    HIS.geometry('300x200')

    L =Label(HIS, text = 'DEFECT HISTORY',font = Font3).pack()
    #History table
    header =['Datetime', 'Item Name','Defect1']
    hwidth = [100,100,80]

    table_history =ttk.Treeview(HIS, columns=header, show = 'headings', height =10)
    table_history.pack()

    #for loop สำหรับตาราง header  
    for hd,hw in zip(header,hwidth):
        table_history.column(hd, width =hw)
        table_history.heading(hd, text =hd)

    #Update from csv
    try:
        with open('transaction.csv', newline ='',encoding='utf-8') as file:
            fr =csv.reader(file) # file reader
            for row in fr:
                table_history.insert('',0,value= row)
    except:
        messagebox.showinfo('No database','No database\nฐานข้อมูลถูกลบไปแล้ว!!!')

    HIS.mainloop()
GUI.bind('<F1>', HistoryWindow)


def DeleteCSV():
    global table_history, HIS
    HIS = Toplevel()
    HIS.geometry('200x250')
    HIS.title('Please input password')

    #messagebox.askyesnocancel('???','Will you delete all data?\nAre you sure?')

    v_password = StringVar()
    E = Entry(HIS, show='*', textvariable=v_password, width = 10, font=Font3)
    E.pack(pady=50)

    def checkPW():

        pw =  v_password.get()
        if pw=='212224236':

            messagebox.showinfo('Password','Valid password\nรหัสผ่านถูกต้อง')
            
            response= messagebox.askyesnocancel('Delete database','Delete all database?\nลบฐานข้อมูลทั้งหมด คุณแน่ใจหรือไม่')
            #print(response)
            #กด Yes คืนค่า True, กด No/Cancell คืนค่า False และเก็บผลไว้ที่ตัวแปร  response
            if response ==True:
                messagebox.showinfo('Delete database','Delete all database...\nกำลังลบฐานข้อมูลทั้งหมด')
                time.sleep(3)
                os.remove('transaction.csv')
            else:
                pass

        else:
            messagebox.showinfo('Password','Invalid password\nรหัสผ่านไม่ถูกต้อง')
            pass

    def show_password():
        if E.cget('show')=='*':
            E.config(show='')
        else:
            E.config(show='*')

    check_button = Checkbutton(HIS, text='show password',font=Font2, command=show_password)
    check_button.place(x=30, y=100)

    check_password = Button(HIS, text='CHECK PASSWORD',font=Font2, command=checkPW)
    check_password.place(x=30,y=130)

    HIS.mainloop()

##########Export database ##############################################
menubar = Menu(GUI)
GUI.config( menu = menubar)

filemenu = Menu(menubar, tearoff=0)
menubar.add_cascade(label ='File', menu = filemenu)


def ExportDatabase():

    #read csv file
    csv = pd.read_csv('transaction.csv',header =None, names=['DateTime', 'Item Name','Defect1']  )

    #create excel writer
    excelWriter = pd.ExcelWriter('/home/manu1-pi1/Desktop/HistoryData/data.xlsx')
    #excelWriter = pd.ExcelWriter('Q:/06) QC1/11_HistoryTBCamera/data.xlsx')
    

    #convert CSV to excel
    csv.to_excel(
        excelWriter,
        index_label='No.',
        freeze_panes= (1,0), #freeze แถวแรก, column ที่ ศูนย์
        sheet_name= ' History data from  CSV'
        )

    messagebox.showinfo('CSV_to_Excel','กำลัง Export ข้อมูล...')
    time.sleep(3)
    #save excel file
    #excelWriter.save()
    excelWriter.close()


#เพิ่มมนูย่อย
filemenu.add_command(label = 'Export', command=ExportDatabase)
filemenu.add_command(label = 'Delete ALL History', command = DeleteCSV)
filemenu.add_command(label = 'Exit', command=GUI.quit)

#####################################################################
'''
fall_signal = False

def trigger(event):
    global fall_signal
    if inPin !=0:
        fall_signal=True
        #dt = datetime.now()
        print('trigger:', inPin)
    else:
        pass
        print('no-trigger')
'''

GPIO.setmode(GPIO.BOARD) #GPIO4
GPIO.setwarnings(False) #บางครั้งมีการใช้ GPIO ตัวเก่าไปแล้ว กัไว้ไม่ให้มี error
relay= 7
#inPin = 11
OutPin=13

#GPIO.setup(inPin, GPIO.IN, pull_up_down=GPIO.PUD_UP)
#GPIO.add_event_detect(inPin, GPIO.FALLING, bouncetime=100)
#GPIO.add_event_callback(inPin, trigger)

GPIO.setup(relay, GPIO.OUT)
GPIO.setup(OutPin, GPIO.OUT)

GPIO.output(relay,True)

GPIO.output(OutPin,True)
time.sleep(1)
GPIO.output(OutPin,False)


cap = None

def Start():
    global cap
    try:

        cap =cv2.VideoCapture(0)
        if v_cameraID.get()==0:
            update_img() #black printed inspection
        else:
            update_img1() #white printed inspection
        
    except:
        messagebox.showwarning('Info', 'No Camera signal')
 
def Stop():
    try:
        cap.release()
        GPIO.output(relay,True)
        GPIO.output(OutPin,False)
    except:
        pass
    
def to_pil(img,label,x,y, H1, H2, W1, W2):
    #img = cv2.resize(img, (w, h))
    img = img[H1:H2,W1:W2]
    image = Image.fromarray(img)
    imgTk = ImageTk.PhotoImage(image)
    label.configure(image=imgTk)
    label.image = imgTk
    label.place(x=x, y=y)

frame=None

kernel = np.ones((5,5),np.uint8)

######Black printing inspection
def update_img():
    global frame,cap,  frame_resize,threshTk, thresh, NG_text, NGPixel, OK_text,OKPixel, area,Outpin

    try:

        thresh_value1 = var1.get()
        thresh_value2 = var7.get()

        A_Min_bar= var2.get()
        H1=  var3.get()
        H2= var4.get()
        W1 = var5.get()
        W2 = var6.get()
        
        ExImg = v_ExImg.get()
        _,frame =cap.read()
        frame=cv2.flip(frame,1)

        frame_resize = cv2.resize(frame,(ExImg*420,ExImg*480))
        frame_ROI = frame_resize[ExImg*H1:ExImg*H2, ExImg*W1:ExImg*W2] # set up ROI

        frame_rgb= cv2.cvtColor(frame_ROI, cv2.COLOR_BGR2RGB)

        #frame =cv2.resize(frame,(H1,W1))
        frame = ImageTk.PhotoImage(Image.fromarray(frame_rgb))
        canvas1.create_image(100,100, image=frame, anchor='center')
        
        Blurred = cv2.GaussianBlur(frame_rgb,(5,5),0)
        frame_gray= cv2.cvtColor(Blurred, cv2.COLOR_BGR2GRAY)
                
        _,threshold = cv2.threshold(frame_gray,thresh_value1,255,cv2.THRESH_BINARY)
        
        thresh = ImageTk.PhotoImage(Image.fromarray(threshold))
        canvas2.create_image(100,100, image=thresh, anchor='center')
        
        image_size = threshold.size
        whitePixels =cv2.countNonZero(threshold)
        blackPixels = image_size - whitePixels
        
        counter = 0

              
        #Calculate and output:
        if blackPixels>A_Min_bar:
            NG_text = 'NG'

            v_status.set('BlackPixel = {} {}'.format(int(blackPixels), NG_text))
            L1.configure(fg='red')
            
            if on_sw==True:
                GPIO.output(relay,False)
                GPIO.output(OutPin,True)
        
                v_defect1 = NG_text 
                AddTransaction()
                #print('relay on')

            else:
                pass
            
        else:
            OK_text = 'OK'
            v_status.set('BlackPixel = {} {}'.format(int(blackPixels), OK_text))
            L1.configure(fg='green')
            GPIO.output(relay,True)
            GPIO.output(OutPin,False)
            #print('relay off')

        GUI.after(30,update_img) #repeat image every 33 fps

    except KeyboardInterrupt:
        GPIO.cleanup()
        messagebox.showwarning('Info','Camera Stop!, No signal')

######white printed inspection
def update_img1():
    global frame,cap,  frame_resize,threshTk, thresh, NG_text, NGPixel, OK_text,OKPixel, area

    try:

        thresh_value1 = var1.get()
        thresh_value2 = var7.get()

        A_Min_bar= var2.get()
        H1=  var3.get()
        H2= var4.get()
        W1 = var5.get()
        W2 = var6.get()
        ExImg = v_ExImg.get()
        
        _,frame =cap.read()
        frame=cv2.flip(frame,1)

        frame_resize = cv2.resize(frame,(ExImg*420,ExImg*480))
        frame_ROI = frame_resize[ExImg*H1:ExImg*H2, ExImg*W1:ExImg*W2] # set up ROI

        frame_rgb= cv2.cvtColor(frame_ROI, cv2.COLOR_BGR2RGB)

        #frame =cv2.resize(frame,(H1,W1))
        frame = ImageTk.PhotoImage(Image.fromarray(frame_rgb))
        canvas1.create_image(100,100, image=frame, anchor='center')
        
        Blurred = cv2.GaussianBlur(frame_rgb,(5,5),0)
        frame_gray= cv2.cvtColor(Blurred, cv2.COLOR_BGR2GRAY)
        
        #print("camID=", v_cameraID.get())
                
        _,threshold = cv2.threshold(frame_gray,thresh_value1,255,cv2.THRESH_BINARY)
        
        thresh = ImageTk.PhotoImage(Image.fromarray(threshold))
        canvas2.create_image(100,100, image=thresh, anchor='center')
        
        image_size = threshold.size
        whitePixels =cv2.countNonZero(threshold)
        blackPixels = image_size - whitePixels
        
        counter = 0

        if whitePixels<A_Min_bar:
            
            NG_text = 'NG'

            v_status.set('WhitePixel = {} {}'.format(int(whitePixels), NG_text))
            L1.configure(fg='red')
            GPIO.output(OutPin,False)
            
            if on_sw==True:
                GPIO.output(relay,False)
                GPIO.output(OutPin,True)
                                
                v_defect1 = NG_text 
                AddTransaction()
                #print('relay on')
            else:
                pass
         
        else:
            OK_text = 'OK'
            v_status.set('WhitePixel = {} {}'.format(int(whitePixels), OK_text))
            L1.configure(fg='green')
            GPIO.output(relay,True)
            GPIO.output(OutPin,False)
            #print('relay off')

    #showframe()
    #update_img()

        GUI.after(30,update_img1) #repeat image every 33 fps

    except KeyboardInterrupt:
        GPIO.cleanup()
        messagebox.showwarning('Info','Camera Stop!, No signal')

######Button & Entry ########
v_itemname=StringVar()

ItemName_Label = tk.Label(frame_1, text='Item Name', font =Font1)
ItemName_Label.place(x=10,y=0)

ItemName_entry = tk.Entry(frame_1, textvariable=v_itemname, width = 15, font =Font2)
ItemName_entry.place(x=10,y=25)

Start_Button = tk.Button(frame_1, text='START',width=5, font=Font2, command = Start)
Start_Button.place(x=0,y=50)

Stop_Button = tk.Button(frame_1, text='STOP',width=5, font=Font2, command =Stop)
Stop_Button.place(x=70,y=50)

#Exit_Button = tk.Button(frame_1, text='EXIT',width=5, font=Font2, fg='red', command = GUI.quit)
#Exit_Button.place(x=140,y=50)


W = 150

var1 = IntVar()
var2 = IntVar()
var3 = IntVar()
var4 = IntVar()
var5 = IntVar()
var6 = IntVar()
var7 = IntVar()


Thresh1 = tk.Scale(frame_2, label="Threshold1", from_=1, to=255, orient=HORIZONTAL, font=Font2, variable=var1)
Thresh1.set(50)
Thresh1.place(x=0, y=0, width=W)

# Thresh2 = tk.Scale(frame_2, label="Threshold2", from_=126, to=255, orient=HORIZONTAL, variable=var7)
# Thresh2.set(255)
# Thresh2.place(x=10, y=70, width=W)

A_Min0 = tk.Scale(frame_2, label="Pixel_Low-limit", from_=10, to=20000, orient=HORIZONTAL,font=Font2, variable=var2)
A_Min0.set(200)
A_Min0.place(x=150, y=0, width=W)


#set up จอจับภาพ

H1 = tk.Scale(frame_3, label="H1", from_=1, to=250, orient=VERTICAL,font=Font3, variable=var3)
H1.set(10)
H1.place(x=10,y=20)

H2 = tk.Scale(frame_3, label="H2", from_=251, to=500, orient=VERTICAL, font=Font3,variable=var4)
H2.set(640)
H2.place(x=10, y=130)

W1 = tk.Scale(frame_3, label="W1", from_=1, to=250, orient=VERTICAL,font=Font3, variable=var5)
W1.set(10)
W1.place(x=90,y=20)

W2 = tk.Scale(frame_3, label="W2", from_=251, to=500, orient=VERTICAL,font=Font3, variable=var6)
W2.set(720)
W2.place(x=90, y=130)

L = tk.Label(frame_3, text='Set Capture Area ', font = Font2)
L.place(x=10,y=0)

v_status = StringVar()
v_status.set('<<<No-Status>>>')
L1 = tk.Label(frame_4,textvariable=v_status, font= Font2)
L1.pack()


###Camera_ID###
option1=['0','1']
tk.Label(frame_3, text= 'Camera\nID', font=Font3).place(x=10, y=240)
v_cameraID =IntVar()
v_cameraID.set(option1[0])
drop1= OptionMenu(frame_3,v_cameraID,*option1)
drop1.place(x=10,y=270)
drop1.config(font=Font3)

###Expand size of image###
option2=['1','2','3']
tk.Label(frame_3, text= 'Expand\nImage', font=Font3).place(x=100, y=240)
v_ExImg =IntVar()
v_ExImg.set(option2[0])
drop2= OptionMenu(frame_3,v_ExImg,*option2)
drop2.place(x=90,y=270)
drop2.config(font=Font3)

##########Radio button ###########

v_rbutton = IntVar()
v_rbutton.set(0)


def clicked(value):
    #print(v_rbutton.get())
    r_label = Label(frame_3, text = value, font=Font3)
    r_label.place(x=150, y=200)


Save_button = Radiobutton(frame_3, text= 'SAVE\nDATA',font=Font3,  variable= v_rbutton, value = 1, command = lambda: clicked(v_rbutton.get()))
Save_button.place( x=150, y=220)

NotSave_button = Radiobutton(frame_3, text= 'NOT\nSAVE\nDATA',font=Font3, variable= v_rbutton, value = 2, fg='red', command = lambda: clicked(v_rbutton.get()))
NotSave_button.place( x=150, y=250)

#print(v_rbutton.get())

# r_label = Label(frame_2, text = v_rbutton.get())
# r_label.place(x=0, y=470)

##toggle button ###
on_sw= False
def switch():
    global on_sw
    if on_sw:
        label_on.config(text='OFF', fg='red')
        button_toggle.config(image=toggle_off)
        on_sw=False
    else:
        label_on.config(text='ON',fg='green')
        button_toggle.config(image=toggle_on)
        on_sw=True

v_toggle = IntVar()
label_on=tk.Label(frame_1, text='OFF',font=Font3, fg='red')
label_on.place(x=140,y=5)
toggle_on=PhotoImage(file='/home/manu1-pi1/Desktop/Python code/toggle-switch-on.png')
toggle_off=PhotoImage(file='/home/manu1-pi1/Desktop/Python code/toggle-switch-off.png')

button_toggle=Button(frame_1,text='OFF',fg='red', image=toggle_off, bd=0, command=switch)
button_toggle.place(x=140,y=30)


GUI.mainloop()
