from customtkinter import *
from PIL import Image
from customtkinter import CTkImage
from tkinter import messagebox

def login():
    if usernameEntry.get() == '' or passwordEntry.get() == '':
        messagebox.showerror('错误', '请输入完整信息')
    elif usernameEntry.get() == 'tl' and passwordEntry.get() == '1234':
        messagebox.showinfo('成功', '登陆成功！')
        root.destroy()
        
        # 导入主界面模块
        import ems_new
        
        # 创建并启动主界面
        new_root = CTk()
        app = ems_new.EmployeeManagementSystem(new_root)
        new_root.mainloop()
    else:
        messagebox.showerror('错误', '信息错误')

root = CTk()
root.geometry('930x478')
root.resizable(0, 0)

root.title('企业员工管理系统 - 登录')
image = CTkImage(Image.open('cover.png'), size=(930, 478))
imageLabel = CTkLabel(root, image=image, text='')
imageLabel.place(x=0, y=0)
headinglabel = CTkLabel(root, text='企业员工管理系统',
                        bg_color='#FBFBFB',
                        font=('Goudy Old Style', 20, 'bold'),
                        text_color='dark blue')
headinglabel.place(x=80, y=100)

usernameEntry = CTkEntry(root, placeholder_text='请输入用户名', width=180)
usernameEntry.place(x=50, y=150)

passwordEntry = CTkEntry(root, placeholder_text='请输入密码', width=180, show='*')
passwordEntry.place(x=50, y=200)

loginButton = CTkButton(root, text='登录', cursor='hand2', command=login)
loginButton.place(x=70, y=250)

root.mainloop()
