#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pdfplumber
import glob
import os
import re
# import pandas as pd
# from tqdm.notebook import tqdm
import tkinter
from tkinter import filedialog, messagebox, StringVar
# import xlwt
from pandas import DataFrame
import openpyxl
# from datetime import date


# In[2]:


def find_2_5(pdf):
    for i in range(len(pdf.pages)):
            page = pdf.pages[i]
            text = str(page.extract_text())
            text = text.split('\n')
            for j in range(len(text)):
                if text[j].lower().startswith('2.5.'):
    #                 print(text[j])
                    next_line = 1

                    try:
                        while len(text[j+next_line]) < 10:
                            next_line += 1
                        if 'ơn vị tính' in text[j+next_line].lower(): # python doesnot recognize the letter 'đ' in đơn vị
                            next_line +=1

                        a = text[j+next_line]
                        return a
                    except:
                        page_next = pdf.pages[i+1]
                        text_next = str(page_next.extract_text())
                        text_next = text_next.split('\n')
                        
#                         text_next = text_next.split('\n')
                        next_line = 0
                        while len(text_next[next_line]) <9:
                            next_line +=1
                        a_next = text_next[next_line]
                        return a_next
#                         print(a_next)


                elif 'không có thông tin' in text[j].lower():
                    b = 'CIC không có thông tin'
                    return b
#                     print(b)

# def find_id_card(text):# Quan update 01/08/2021 for new format of CIC report, where "số chứng minh nhân dân" inline with the number
#     text = str(text).split('\n')
#     IDcard= []
#     for row in text:
#         if 'số chứng minh' in row.lower() and  re.findall('[0-9]*$',row) != [] and IDcard != [] :
#             try:
#                 id_card = re.findall('[0-9]*$',row)
#     #             print(id_card)
# #                 return id_card[0].strip()
#                 IDcard.append(id_card[0].strip())
#             except:
# #                 return 'can not map ID'
#                 continue
#         elif 'CMT' in row and 'chính xác' in row.lower():
#             try:
#                 id_card = re.findall('.*?([0-9]+)$',row)
# #                 return id_card[0].strip()
#                 IDcard.append(id_card[0].strip())
#             except:
# #                 return 'wrong format/NO CIC'
#                 continue
#     try:
#         return IDcard[0]
#     except:
#         return 'cannot map ID'
    
def find_id_card(pdf):# Quan update 01/08/2021 for new format of CIC report, where "số chứng minh nhân dân" inline with the number

    page = pdf.pages[0] #client id in first page
    text = str(page.extract_text())
    
    text = str(text).split('\n')
    IDcard= []
    for row in text:
        if 'số chứng minh' in row.lower() and  re.findall('[0-9]*$',row) != []:
            try:
                id_card = re.findall('[0-9]*$',row)
                IDcard.append(id_card[0].strip())
            except:
                continue
        elif 'CMT' in row and 'chính xác' in row.lower():
            try:
                id_card = re.findall('.*?([0-9]+)$',row)
                IDcard.append(id_card[0].strip())
            except:
                continue
    try:
        return IDcard[0]
    except:
        return 'cannot map ID'
                

# def find_client_name(text):
#     text = str(text).split('\n') 
#     for row in text:
#         if row.lower().startswith('tên') and 'khách' in row.lower() and ':' in row.lower():
#             try:
#                 client_name = re.findall('(?<=: )[^\]]+',row)[0]
#                 return client_name.strip()
#             except:
#                 return 'Can NOT find name'
#         elif row.lower().startswith('tên') and 'khách' in row.lower() and ':' not in row.lower():
#             try:
                
#                 client_name = re.findall('(?<=Tên khách hàng).*$',row)[0]
# #                 err = 'wrong format file'
# #                 return err
#                 return client_name.strip()
#             except:
#                 return 'Can NOT find name'
        
def find_client_name(pdf):
    page = pdf.pages[0] #client name in first page
    text = str(page.extract_text())
    text = str(text).split('\n')
    
    for row in text:
        if row.lower().startswith('tên') and 'khách' in row.lower() and ':' in row.lower():
            try:
                client_name = re.findall('(?<=: )[^\]]+',row)[0]
                return client_name.strip()
            except:
                return 'Can NOT find name'
        elif row.lower().startswith('tên') and 'khách' in row.lower() and ':' not in row.lower():
            try:
                
                client_name = re.findall('(?<=Tên khách hàng).*$',row)[0]
#                 err = 'wrong format file'
#                 return err
                return client_name.strip()
            except:
                return 'Can NOT find name'


# out_put = {'RA_NAME':[], 'RA_ID':[], 'history_3y':[]}
# data = pd.DataFrame(out_put)

def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global saveLocation
    saveLocation = filedialog.askdirectory() +'/'
    entryPath.delete(0,'end')
    entryPath.insert(0,saveLocation)
    #print(saveLocation)


# In[3]:


def read_pdf():
#     today = date.today()
    out_put = []
    total_file = len([x for x in glob.glob(os.path.join(saveLocation, '*.pdf'))])
    cnt = 0
#     for filename in tqdm(glob.glob(os.path.join(saveLocation, '*.pdf'))):
    for filename in glob.glob(os.path.join(saveLocation, '*.pdf')):
        cnt +=1
        short_name = filename.replace(saveLocation[:-1] + '\\', '')
        statusText.set(str(cnt) + '/' + str(total_file) + ':Processing of file #'+ str(short_name))
        status.update()
        print(filename)
        pdf = pdfplumber.open(filename)
        history = find_2_5(pdf)   # Quan fix bug 2.5 next page
        
        # Quan fix 01/08/2021 for new format in CIC: new function
        client_id = find_id_card(pdf) 
        client_name = find_client_name(pdf)
        
        if client_name != None:
            name = client_name
        else:
            name = None

        if client_id != None:
            cid = client_id
        else:
            cid = None

        if history != None:
            if 'không có nợ xấu' in history:
                his =  history
            elif 'không có thông tin' in history:
                his = 'CIC không có thông tin'
            else:
                his = 'NỢ XẤU: vui lòng kiểm tra file PDF'
        else:
            his = 'wrong format/NO CIC'

#         new_row = {'RA_NAME':name, 'RA_ID':cid, 'history_3y':his}
        new_row = {'RA_NAME':name, 'RA_ID':cid, 'history_3y':his, 'filename': short_name}
    #     short_name = os.path.splitext(short_name)[0] #split extension to take only filename
        out_put.append(new_row)

    data = DataFrame(out_put, index = None)
    try:
#         data.to_excel(saveLocation + 'CIC_RESULT_' +str(total_file) + 'files' + '_'+ str(today)+ '.xlsx', encoding = 'utf-8', index = None)
        data.to_excel(saveLocation + 'CIC_RESULT_' +str(total_file) +'_'+ 'files' +'.xlsx', encoding = 'utf-8', index = None)
        messagebox.showinfo(title = "Completed", message = "All file read! Please check file CIC_RESULT.xlsx in your directory!!!")
    except:
        messagebox.showinfo(title = "Error opening file", message = "Please close file CIC_RESULT.xlsx then run again")
    
#     messagebox.showinfo("Completed","All file read! Please check file CIC_RESULT.xlsx in your directory!!!")
    


# In[4]:


# import openpyxl


# In[5]:


root = tkinter.Tk()
root.geometry("500x350")
root.title("Read CIC file V2")

# root.configure(background='bisque3')

# label1 = tkinter.Label(root, text="skp_client")
# label1.place(x=0, y=30)

# inputBox = tkinter.Text(root, height=20, width=20)
# inputBox.place(x=0, y=50)

# Oracle Username and password
# labelUser = tkinter.Label(root, text="Oracle username:")
# labelUser.place(x=40, y=80)

# inputUser = tkinter.Entry(root, width=40)
# inputUser.place(x=140, y=80)

# labelPassword = tkinter.Label(root, text="Oracle password:")
# labelPassword.place(x=40, y=110)

# inputPassword = tkinter.Entry(root, show="*", width=40)
# inputPassword.place(x=140, y=110)




# # BSL/Cabinet Username and password  -- adding 31/12/2019
# BSL_labelUser = tkinter.Label(root, text="BSL username :")
# BSL_labelUser.place(x=40, y=150)

# BSL_User = tkinter.Entry(root, width=35)
# BSL_User.place(x=140, y=150)

# BSL_labelPassword = tkinter.Label(root, text="BSL password :")
# BSL_labelPassword.place(x=40, y=180)

# BSL_Password = tkinter.Entry(root, show="*", width=35)
# BSL_Password.place(x=140, y=180)



# Save directory
label2 = tkinter.Label(root, text="CIC pdf directory:")
label2.place(x=40, y=70)

entryPath = tkinter.Entry(root, width=40)
global saveLocation
saveLocation = os.path.expanduser('~/Documents/')
entryPath.insert(0,saveLocation)
entryPath.place(x=140, y=70)

btnBrowse = tkinter.Button(root, text = "Browse...", command = browse_button)
btnBrowse.place(x=400, y=70)


btnDownload = tkinter.Button(root, text = "Read PDF CIC", height=3, width=30, command = read_pdf)
btnDownload.place(x=150, y=150)


statusText = StringVar()
status = tkinter.Label(root, textvariable=statusText, borderwidth=2, relief="sunken") 
status.pack(side='bottom') 
# statusText = StringVar()
# status = tkinter.Label(root, textvariable=statusText, borderwidth=2, relief="sunken") 
# status.pack(side='bottom') 

# loginFile = os.path.expanduser('~/Documents/') + 'loginfile.pickle'
# if os.path.exists(loginFile):
#     fileObject = open(loginFile,'rb')  
#     infoList = pickle.load(fileObject)
#     inputUser.insert(0,infoList[0])
#     inputPassword.insert(0,infoList[1])
#     fileObject.close()


root.mainloop()


# In[ ]:




