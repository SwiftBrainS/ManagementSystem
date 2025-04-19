import database
import os
import sys
from tkinter import messagebox
import tkinter as tk

def main():
    required_files = ['cover.png', 'bg.png']
    missing_files = [f for f in required_files if not os.path.exists(f)]
    
    if missing_files:
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        messagebox.showerror('错误', f'找不到必要的文件: {", ".join(missing_files)}')
        root.destroy()
        sys.exit(1)
    
    # 初始化数据库
    try:
        # 连接数据库
        database.connect_database()
        
        # 检查是否已存在数据，如果不存在则添加样例数据
        mycursor = database.mycursor
        mycursor.execute("SELECT COUNT(*) FROM Department")
        dept_count = mycursor.fetchone()[0]
        
        if dept_count == 0:
            database.insert_sample_data()
        
        # 启动登录界面
        import login
    except Exception as e:
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        messagebox.showerror('错误', f'初始化数据库时出错: {str(e)}')
        root.destroy()
        sys.exit(1)

if __name__ == "__main__":
    main() 