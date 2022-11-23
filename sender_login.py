import tkinter as tk
import webbrowser
import smtplib
import sender_tkGUI

# 取得 input
def getLoginInput(event=None):
    account = enry1.get()
    password = enry2.get()
    loginSMTP(account,password)

# login smtp
def loginSMTP(account,password):
    try:
        with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 設定SMTP伺服器
            try:
                smtp.ehlo()  # 驗證SMTP伺服器
                smtp.starttls()  # 建立加密傳輸
                smtp.login(account,password)  # 登入寄件者gmail
                tk.messagebox.showinfo("^_^", '登入成功，選擇檔案\nLogin successful. Please select a file')
                root.destroy()
                sender_tkGUI.exe_choose_file(account,password)
            except Exception as e:
                tk.messagebox.showinfo("*_*", '信箱或SMTP輸入錯誤，請重新輸入\nWrong User Gamil or Password')
    except:
        tk.messagebox.showinfo("*_*", '沒有網路連線')

# 超連結
def callback(url):
    webbrowser.open_new(url)

if __name__ == "__main__": 
    root = tk.Tk()
    root.title('Personalized Mass Email Sender Login')
    root.configure(background='#e9dfed')
    root.geometry('+360+200')
    root.minsize(520, 250)
    root.maxsize(520, 250)
    # root.iconbitmap("myIcon.ico")

    # Label
    label0 = tk.Label(text="請輸入寄信者 gmail 及 SMTP               \nPlease enter gamil address and SMTP", fg='black', bg="#e9dfed", font=("Helvetica",'12'))
    label1 = tk.Label(text="User gmail", fg='black', bg="#e9dfed", font=("Helvetica",'12'))
    label2 = tk.Label(text="SMTP", fg='black', bg="#e9dfed", font=("Helvetica",'12'))

    label0.grid(column=0, row=0, columnspan=2, pady=20, padx=5, sticky='w')
    label1.grid(column=0, row=1, padx=5)
    label2.grid(column=0, row=2, padx=5)

    # Entry
    enry1 = tk.Entry(width=45, font=("Helvetica",'12'))
    enry2 = tk.Entry(width=45, font=("Helvetica",'12'), show='*')

    enry1.grid(column=1, row=1, padx=10, pady=10)
    enry2.grid(column=1, row=2, padx=10, pady=10)
    enry1.bind('<Return>', getLoginInput)
    enry2.bind('<Return>', getLoginInput)

    # Button
    button1 = tk.Button(text='Login',width=20, bg="white", font=("Helvetica",'12','bold'), command=getLoginInput)
    button1.grid(column=1, row=3, pady=2, padx=15,sticky='e')

    # Lable, bg="#e9dfed"
    label_smtp = tk.Label(text="*SMTP 為 gmail 存取全密碼，取得SMTP教學網址:", fg='black', bg="#e9dfed")
    link1 = tk.Label(root, text="https://www.webdesigntooler.com/google-smtp-send-mail", fg="blue", bg="#e9dfed", cursor="hand2")
    link1.bind("<Button-1>", lambda e: callback("https://www.webdesigntooler.com/google-smtp-send-mail"))

    label_smtp.grid(column=0, row=4,columnspan=2, padx=10, sticky='w')
    link1.grid(column=0, row=5,columnspan=2, padx=10,pady=5, sticky='w')

    root.mainloop()
