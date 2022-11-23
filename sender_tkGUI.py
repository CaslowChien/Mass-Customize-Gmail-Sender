import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import filedialog
import sender_send
from tkinter import *
from PIL import ImageTk, Image
import base64
from pic2str import icon
from io import BytesIO

# 大區塊分割
def define_layout(obj, cols=1, rows=1):
    
    def method(trg, col, row):
        
        for c in range(cols):    
            trg.columnconfigure(c, weight=1)
        for r in range(rows):
            trg.rowconfigure(r, weight=1)

    if type(obj)==list:        
        [ method(trg, cols, rows) for trg in obj ]
    else:
        trg = obj
        method(trg, cols, rows)

def exe_choose_file(user_account, user_password):
    def loadFile():
        file_path = filedialog.askopenfilename(title = "Select file",filetypes = (("CSV Files","*.csv"),("Excel files", "*.xlsx")))
        if file_path[-5:]==".xlsx":
            df = pd.read_excel(file_path)
        elif file_path[-4:]=='.csv':
            df = pd.read_csv(file_path)
        window0.destroy()
        exe(df, user_account, user_password)
    window0 = tk.Tk()
    window0.title('Select file')
    window0.minsize(450, 250)
    window0.maxsize(450, 250)
    window0.geometry('+400+150')
    window0.config(bg="#e9dfed")
    # window0.iconbitmap("myIcon.ico")

    lb1 = tk.Label(window0, text='請選擇 csv 或 xlsx 檔\nPlease select a csv or xlsx file', bg="#e9dfed", font=("Helvetica",'12'))
    lb1.pack(anchor=tk.N, pady=30)

    # 選取檔案 (command 為選取檔案並把檔案中的 col 名加入到 div1 裡)
    bt1 = tk.Button(window0, height=4, width=17, text='選取檔案\nSelect a file', command=loadFile, font=("Helvetica",'10'))
    # 按下後會取得 df
    bt1.pack(pady=10)

    window0.mainloop()
  
def exe(df, account, smtp):
    window = tk.Tk()
    window.title('Personalized Mass Email Sender')
    window.geometry('900x500+160+70')
    window.minsize(900, 500)
    # window.iconbitmap("myIcon.ico")
    align_mode = 'nswe'
        
    """Frame (版面配置)"""
    div_size = 200
    img_size = 50
    div0 = tk.Frame(window,  width=900 , height=30, bg='#b5a4bd')
    div1 = tk.Frame(window,  width=400 , height=500)
    div2 = tk.Frame(window,  width=500 , height=200)
    div3 = tk.Frame(window,  width=500 , height=300, bg='#d2c5d9')

    div0.grid(column=0, row=0, columnspan=2, sticky=align_mode)
    div1.grid(column=0, row=1, rowspan=2, sticky=align_mode)
    div2.grid(column=1, row=1)
    div3.grid(column=1, row=2, sticky=align_mode)

    define_layout(window, cols=2, rows=3)
    define_layout([div0, div1, div2, div3])

    '''---div0--- Logo列'''
    byte_data = base64.b64decode(icon)
    image_data = BytesIO(byte_data)
    image = Image.open(image_data)

    # image=Image.open('logo+mark+background.png')
    img=image.resize((350,30),Image.ANTIALIAS)
    img=ImageTk.PhotoImage(img)
    panel = Label(div0, image = img, height=30,width=350, bg="#b5a4bd")
    panel.grid(column=0, row=0)

    '''---div1--- 檔案col顯示'''

    replace_lst=[]
    def cb(event):
        index = div1_lb.curselection()
        div3_text.insert('insert',"<<{", 'iterating')
        div3_text.insert(tk.INSERT,div1_lb.get(index), 'iterating')
        div3_text.insert('insert', "}>>", 'iterating')
        if ('<<{'+div1_lb.get(index)+"}>>") not in replace_lst:
            replace_lst.append('<<{'+div1_lb.get(index)+"}>>")
    
    # 右邊卷軸
    div1_sb1 = tk.Scrollbar(div1)
    div1_sb1.grid(column=1, sticky=align_mode)
    # 下面的卷軸
    div1_sb2= tk.Scrollbar(div1, orient=tk.HORIZONTAL)
    div1_sb2.grid(column=0, sticky=align_mode)
    # 左邊 col 的 Listbox 
    div1_lb = tk.Listbox(div1, yscrollcommand=div1_sb1.set, xscrollcommand=div1_sb2.set,\
        selectmode=tk.SINGLE, bg="#e9dfed", font=("Helvetica",'10'), activestyle="dotbox",\
        exportselection=0, selectborderwidth=5, selectbackground="#483d4d",\
        selectforeground='white', relief="sunken")
    div1_lb.grid(column=0, row=0, sticky=align_mode)
    div1_lb.bind('<<ListboxSelect>>', cb)

    # 結合右邊(垂直)卷軸
    div1_sb1.config(command=div1_lb.yview)
    # 結合下面(水平)捲軸
    div1_sb2.config(command=div1_lb.xview)

    # 把 csv 的 columns 放到 listbox 裡
    for col in list(df.columns):
        div1_lb.insert(tk.END, col)

    '''---div2--- 寄信資訊'''
    # label (from, to ,subject)
    div2_title1 = tk.Label(div2, text="From", fg='black')
    div2_title2 = tk.Label(div2, text="To", fg='black')
    div2_title3 = tk.Label(div2, text="Subject", fg='black')
    div2_title4 = tk.Label(div2, text=account, fg='black')

    div2_title1.grid(column=0, row=0, sticky=align_mode)
    div2_title2.grid(column=0, row=1, sticky=align_mode)
    div2_title3.grid(column=0, row=2, sticky=align_mode)
    div2_title4.grid(column=1, row=0, sticky=align_mode)

    # optionlist 選擇收件者
    optionList = [df.dtypes.index[i] for i in range(len(df.dtypes)) if df.dtypes[i] =='object']
    var= tk.StringVar()
    div2_om1 = ttk.OptionMenu(div2, var, '選擇收件人欄位', *optionList)
    div2_om1.grid(column=1, row=1, padx=10, pady=10)
    
    # entry 填入信件標題subject
    div2_entry1 = ttk.Entry(div2, width=40)
    div2_entry1.grid(column=1, row=2, padx=10, sticky=align_mode)
    

    '''---div3--- 信件內容'''
    # Scrollbar
    div3_sb = tk.Scrollbar(div3)
    div3_sb.grid(column=3, row=0, sticky=align_mode)

    # text 郵件內容文字框
    div3_text = tk.Text(div3, yscrollcommand=div3_sb.set, undo=True, autoseparators=True, maxundo=-1)
    div3_text.grid(column=0, row=0, columnspan=2, sticky=align_mode)
    div3_text.tag_config('iterating', foreground="#a3466e")
    
    # 結合 Scrollbar 和 text
    div3_sb.config(command=div3_text.yview)


    #######-div3 寄送和預覽

    # button 預覽第一篇視窗
    def preview():
        show = True
        try: # 得到剛剛使用者輸入的標題和接收者
            subject = div2_entry1.get() 
            recipients = df[var.get()]
            raw_content = div3_text.get("1.0","end")
        except:
            show = False
            tk.messagebox.showinfo("*_*", '未選擇收件人欄位')
        
        if show == True:
            # Frame 預覽視窗
            window_previes = tk.Tk()
            window_previes.title('Preview')
            window_previes.geometry('500x550+150+30')
            window_previes.minsize(580, 500)  # 這裡應該要用 .resizable(0, 0) 比較好，但是已經弄了就算了
            window_previes.maxsize(580, 500)
            # window_previes.iconbitmap("myIcon.ico")
            
            div0_preview = tk.Frame(window_previes,  width=550 , height=50, bg='#d2c5d9')
            div1_preview = tk.Frame(window_previes,  width=550 , height=450)
            div0_preview.grid(column=0, row=0, sticky=align_mode)
            div1_preview.grid(column=0, row=1, sticky=align_mode)
            define_layout(window_previes, cols=0, rows=2)
            define_layout([div0_preview, div1_preview])

            # lable 資訊欄
            div0_lb1 = tk.Label(div0_preview, text='From:', bg='#d2c5d9', fg='black')
            div0_lb2 = tk.Label(div0_preview, text='To:', bg='#d2c5d9', fg='black')
            div0_lb3 = tk.Label(div0_preview, text='Subject:', bg='#d2c5d9', fg='black')
            div0_lb1.grid(column=0, row=0, sticky='w', padx=50)
            div0_lb2.grid(column=0, row=1, sticky='w', padx=50)
            div0_lb3.grid(column=0, row=2, sticky='w', padx=50)

            div0_lb4 = tk.Label(div0_preview, text=account, bg='#d2c5d9', fg='black')
            div0_lb6 = tk.Label(div0_preview, text=subject, bg='#d2c5d9', fg='black')
            div0_lb4.grid(column=1, row=0, sticky='w', padx=50)
            div0_lb6.grid(column=1, row=2, sticky='w', padx=50)

            # combobox 選擇預覽的 gmail 收件人
            def combobox_selected(event): # 改變選擇後預覽變化
                content = raw_content[:]
                for i in replace_lst:
                    try:
                        content = content.replace(i,df[i[3:-3]][div0_om1.current()])
                    except:
                        content = content.replace(i,"None")
                div1_text.delete("1.0","end")
                div1_text.insert("insert", content) 

            previewList = [i for i in recipients]
            div0_om1 = ttk.Combobox(div0_preview, values=previewList)
            div0_om1.grid(column=1, row=1, sticky='w', padx=50, pady=10)

            div0_om1.current(0) # 預設為第一個信箱
            div0_om1.bind('<<ComboboxSelected>>', combobox_selected)
            define_layout(div0_preview, cols=2, rows=3)
            define_layout([div0_lb1, div0_lb2, div0_lb3, div0_lb4, div0_om1, div0_lb6])

            # scrollbar & text 內文欄
            div1_sb = tk.Scrollbar(div1_preview, orient='vertical')
            div1_sb.grid(column=1, row=0, sticky=align_mode)

            div1_text = tk.Text(div1_preview, yscrollcommand=div1_sb.set)
            div1_text.grid(column=0, row=0, sticky=align_mode)

            div1_sb.config(command=div1_text.yview)

            # 預設內文為第一個 gmail (df[i[3:-3]][0])
            content = raw_content[:]
            for i in replace_lst:
                try:
                    content = content.replace(i,df[i[3:-3]][0])
                except:
                    content = content.replace(i,"None")  
                
            # 插入剛剛使用者輸入的內文
            div1_text.insert("insert", content) 

            # 執行此視窗
            window_previes.mainloop()
            print(content)
    
    # 預覽第一篇
    div3_bt1 = tk.Button(div3, command=preview, width=15, text='預覽', bg='#efe9f2', font=("Helvetica",'10'), activebackground="#87758f", activeforeground='white')
    div3_bt1.grid(column=0, row=1, pady=5, padx=5, sticky='w')

    # 寄出信後的通知
    def show_result(sucess_lst, failed_lst):
        # frame
        window_show = tk.Tk()
        window_show.title('Result')
        window_show.geometry('600x370+300+90')
        window_show.resizable(0, 0)
        # window_show.iconbitmap("myIcon.ico")
        
        div0_show = tk.Frame(window_show,  width=300 , height=400, bg='#d2c5d9')
        div1_show = tk.Frame(window_show,  width=300 , height=400)
        div0_show.grid(column=0, row=0)
        div1_show.grid(column=1, row=0)
        define_layout(window_show, cols=2, rows=0)
        define_layout([div0_show, div1_show])

        # div0
        div0_show_sb = tk.Scrollbar(div0_show, orient='vertical')
        div0_show_lb = tk.Label(div0_show, text='信件成功寄送\nEmail Sent Successfully:', bg='#d2c5d9', font=("Helvetica", '12', 'bold'))
        div0_show_text = tk.Text(div0_show, yscrollcommand=div0_show_sb.set, font=("Helvetica",'10'))

        div0_show_sb.grid(column=2, row=1, sticky=align_mode)
        div0_show_lb.grid(column=0, row=0, columnspan=2)
        div0_show_text.grid(column=0, row=1, ipadx=15)
        div0_show_sb.config(command=div0_show_text.yview)
        define_layout(div0_show, cols=2, rows=2)
        define_layout([div0_show_lb, div0_show_sb, div0_show_text])

        for i in sucess_lst:
            div0_show_text.insert("end", '  '+i+'\n') 
        div0_show_text.configure(state='disabled')

        # div1
        div1_show_sb = tk.Scrollbar(div1_show, orient='vertical')
        div1_show_lb = tk.Label(div1_show, text='信件寄送失敗\nEmail Sent Failed:', font=("Helvetica", '12', 'bold'))
        div1_show_text = tk.Text(div1_show, yscrollcommand=div1_show_sb.set, font=("Helvetica",'10'))
        
        div1_show_sb.grid(column=2, row=1, sticky=align_mode)
        div1_show_lb.grid(column=0, row=0, columnspan=2)
        div1_show_text.grid(column=0, row=1, ipadx=15)
        div1_show_sb.config(command=div1_show_text.yview)

        define_layout(div1_show, cols=2, rows=2)
        define_layout([div1_show_lb, div1_show_sb, div1_show_text])

        for i in failed_lst:
            div1_show_text.insert("end", '  '+str(i)+'\n') 
        div1_show_text.configure(state='disabled')

        window_show.mainloop()

    # 寄出信
    def send_all_mail():
        show = True
        try: # 得到剛剛使用者輸入的標題和接收者
            subject = div2_entry1.get() 
            recipients = df[var.get()]
            raw_content = div3_text.get("1.0","end")
        except:
            show = False
            tk.messagebox.showinfo("*_*", '未選擇收件人欄位')
        
        if show == True:
            question = tk.messagebox.askyesno(message='確認寄出所有信件?')
            if question == True:
                lst_ok, lst_not_ok = [], []
                for i in range(len(recipients)):
                    content = raw_content[:]
                    for j in replace_lst:
                        content = content.replace(j, str(df[j[3:-3]][i]))
                    re = sender_send.send(subject, account, smtp, recipients[i], content)
                    if re == True: lst_ok.append(recipients[i])
                    elif re == False: lst_not_ok.append(recipients[i])
                show_result(lst_ok, lst_not_ok)
            elif question == False:
                pass
        
    # 全部寄送
    div3_bt2 = tk.Button(div3, command=send_all_mail, width=15, text='全部寄送', bg='#efe9f2', fg='black', font=("Helvetica",'10','bold'), activebackground="#87758f", activeforeground='white')
    div3_bt2.grid(column=1, row=1,padx=150, pady=5)

    window.mainloop()
    
# exe(pd.read_csv("紀念品預購表單test.csv"), "cswyayayaya@gmail.com", "qgogdsvsdmpbixji")
