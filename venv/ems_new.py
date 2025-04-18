from customtkinter import *
from PIL import Image
from customtkinter import CTkImage
from tkinter import ttk, messagebox, VERTICAL, HORIZONTAL
import database
import datetime

# 全局变量
current_tab = None

# 主应用程序
class EmployeeManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1200x700+50+50')
        self.root.resizable(0, 0)
        self.root.title('企业员工管理系统')
        self.root.configure(fg_color='#161C30')
        
        # 加载顶部图片
        self.header_image = CTkImage(Image.open('bg.png'), size=(1200, 158))
        self.header_label = CTkLabel(root, image=self.header_image, text='')
        self.header_label.grid(row=0, column=0, sticky='ew')  # 添加sticky='ew'使图片水平居中
        
        # 创建选项卡
        self.create_tabs()
        
        # 初始化显示员工管理选项卡
        self.show_employee_tab()
        
    def create_tabs(self):
        # 创建选项卡框架
        self.tab_frame = CTkFrame(self.root, fg_color='#161C30')
        self.tab_frame.grid(row=1, column=0, padx=10, pady=10, sticky='ew')
        
        # 添加选项卡按钮
        self.employee_btn = CTkButton(self.tab_frame, text='员工管理', command=self.show_employee_tab, 
                                      width=150, font=('arial', 16, 'bold'))
        self.employee_btn.grid(row=0, column=0, padx=5, pady=5)
        
        self.department_btn = CTkButton(self.tab_frame, text='部门管理', command=self.show_department_tab, 
                                        width=150, font=('arial', 16, 'bold'))
        self.department_btn.grid(row=0, column=1, padx=5, pady=5)
        
        self.position_btn = CTkButton(self.tab_frame, text='职位管理', command=self.show_position_tab, 
                                      width=150, font=('arial', 16, 'bold'))
        self.position_btn.grid(row=0, column=2, padx=5, pady=5)
        
        self.project_btn = CTkButton(self.tab_frame, text='项目管理', command=self.show_project_tab, 
                                     width=150, font=('arial', 16, 'bold'))
        self.project_btn.grid(row=0, column=3, padx=5, pady=5)
        
        self.participation_btn = CTkButton(self.tab_frame, text='项目参与', command=self.show_participation_tab, 
                                          width=150, font=('arial', 16, 'bold'))
        self.participation_btn.grid(row=0, column=4, padx=5, pady=5)
        
        self.stats_btn = CTkButton(self.tab_frame, text='统计分析', command=self.show_stats_tab, 
                                  width=150, font=('arial', 16, 'bold'))
        self.stats_btn.grid(row=0, column=5, padx=5, pady=5)
        
        # 创建内容框架
        self.content_frame = CTkFrame(self.root, fg_color='#161C30')
        self.content_frame.grid(row=2, column=0, padx=10, pady=10, sticky='nsew')
        
    # 员工管理选项卡
    def show_employee_tab(self):
        global current_tab
        
        # 清空内容框架
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        current_tab = "employee"
        
        # 创建表单框架和表格框架
        self.employee_form_frame = CTkFrame(self.content_frame, fg_color='#161C30')
        self.employee_form_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
        
        self.employee_table_frame = CTkFrame(self.content_frame)
        self.employee_table_frame.grid(row=0, column=1, padx=10, pady=10, sticky='ne')
        
        # 创建员工表单
        self.create_employee_form()
        
        # 创建员工表格
        self.create_employee_table()
        
        # 加载员工数据
        self.load_employee_data()
        
    def create_employee_form(self):
        # 员工ID
        CTkLabel(self.employee_form_frame, text='员工ID', font=('arial', 16, 'bold'), text_color='white').grid(row=0, column=0, padx=20, pady=15, sticky='w')
        self.employee_id_entry = CTkEntry(self.employee_form_frame, font=('arial', 14), width=180)
        self.employee_id_entry.grid(row=0, column=1)
        
        # 姓名
        CTkLabel(self.employee_form_frame, text='姓名', font=('arial', 16, 'bold'), text_color='white').grid(row=1, column=0, padx=20, pady=15, sticky='w')
        self.employee_name_entry = CTkEntry(self.employee_form_frame, font=('arial', 14), width=180)
        self.employee_name_entry.grid(row=1, column=1)
        
        # 性别
        CTkLabel(self.employee_form_frame, text='性别', font=('arial', 16, 'bold'), text_color='white').grid(row=2, column=0, padx=20, pady=15, sticky='w')
        gender_options = ['男', '女']
        self.employee_gender_box = CTkComboBox(self.employee_form_frame, values=gender_options, font=('arial', 14), width=180, state='readonly')
        self.employee_gender_box.grid(row=2, column=1)
        self.employee_gender_box.set(gender_options[0])
        
        # 出生日期
        CTkLabel(self.employee_form_frame, text='出生日期', font=('arial', 16, 'bold'), text_color='white').grid(row=3, column=0, padx=20, pady=15, sticky='w')
        self.employee_birth_entry = CTkEntry(self.employee_form_frame, font=('arial', 14), width=180, placeholder_text='YYYY-MM-DD')
        self.employee_birth_entry.grid(row=3, column=1)
        
        # 联系方式
        CTkLabel(self.employee_form_frame, text='联系方式', font=('arial', 16, 'bold'), text_color='white').grid(row=4, column=0, padx=20, pady=15, sticky='w')
        self.employee_contact_entry = CTkEntry(self.employee_form_frame, font=('arial', 14), width=180)
        self.employee_contact_entry.grid(row=4, column=1)
        
        # 入职日期
        CTkLabel(self.employee_form_frame, text='入职日期', font=('arial', 16, 'bold'), text_color='white').grid(row=5, column=0, padx=20, pady=15, sticky='w')
        self.employee_hire_entry = CTkEntry(self.employee_form_frame, font=('arial', 14), width=180, placeholder_text='YYYY-MM-DD')
        self.employee_hire_entry.grid(row=5, column=1)
        
        # 部门
        CTkLabel(self.employee_form_frame, text='部门', font=('arial', 16, 'bold'), text_color='white').grid(row=0, column=2, padx=20, pady=15, sticky='w')
        self.load_departments()
        self.employee_dept_box = CTkComboBox(self.employee_form_frame, values=self.department_options, font=('arial', 14), width=180, state='readonly')
        self.employee_dept_box.grid(row=0, column=3)
        if self.department_options:
            self.employee_dept_box.set(self.department_options[0])
        
        # 职位
        CTkLabel(self.employee_form_frame, text='职位', font=('arial', 16, 'bold'), text_color='white').grid(row=1, column=2, padx=20, pady=15, sticky='w')
        self.load_positions()
        self.employee_position_box = CTkComboBox(self.employee_form_frame, values=self.position_options, font=('arial', 14), width=180, state='readonly')
        self.employee_position_box.grid(row=1, column=3)
        if self.position_options:
            self.employee_position_box.set(self.position_options[0])
        
        # 按钮框架
        button_frame = CTkFrame(self.employee_form_frame, fg_color='#161C30')
        button_frame.grid(row=6, column=0, columnspan=4, pady=10)
        
        # 新建按钮
        new_btn = CTkButton(button_frame, text='新建', font=('arial', 14, 'bold'), width=100, command=self.clear_employee_form)
        new_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # 添加按钮
        add_btn = CTkButton(button_frame, text='添加', font=('arial', 14, 'bold'), width=100, command=self.add_employee)
        add_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # 更新按钮
        update_btn = CTkButton(button_frame, text='更新', font=('arial', 14, 'bold'), width=100, command=self.update_employee)
        update_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # 删除按钮
        delete_btn = CTkButton(button_frame, text='删除', font=('arial', 14, 'bold'), width=100, command=self.delete_employee)
        delete_btn.grid(row=0, column=3, padx=5, pady=5)
        
    def create_employee_table(self):
        # 搜索框架
        search_frame = CTkFrame(self.employee_table_frame)
        search_frame.grid(row=0, column=0, pady=5, sticky='ew')
        
        search_options = ['EmployeeID', 'Name', 'Department', 'Position']
        self.employee_search_box = CTkComboBox(search_frame, values=search_options, state='readonly', width=120)
        self.employee_search_box.grid(row=0, column=0, padx=5, pady=5)
        self.employee_search_box.set('搜索条件')
        
        self.employee_search_entry = CTkEntry(search_frame, width=200)
        self.employee_search_entry.grid(row=0, column=1, padx=5, pady=5)
        
        search_btn = CTkButton(search_frame, text='搜索', width=80, command=self.search_employee)
        search_btn.grid(row=0, column=2, padx=5, pady=5)
        
        show_all_btn = CTkButton(search_frame, text='显示全部', width=80, command=self.load_employee_data)
        show_all_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 员工表格
        columns = ('员工ID', '姓名', '性别', '出生日期', '联系方式', '入职日期', '部门', '职位')
        
        self.employee_tree = ttk.Treeview(self.employee_table_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.employee_tree.heading(col, text=col)
            if col in ['联系方式', '入职日期', '出生日期']:
                self.employee_tree.column(col, width=140)
            elif col == '职位':
                self.employee_tree.column(col, width=120)
            else:
                self.employee_tree.column(col, width=80)
        
        self.employee_tree.grid(row=1, column=0, sticky='nsew')
        
        # 滚动条
        scrollbar = ttk.Scrollbar(self.employee_table_frame, orient=VERTICAL, command=self.employee_tree.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        self.employee_tree.configure(yscrollcommand=scrollbar.set)
        
        # 添加水平滚动条
        hscrollbar = ttk.Scrollbar(self.employee_table_frame, orient=HORIZONTAL, command=self.employee_tree.xview)
        hscrollbar.grid(row=2, column=0, sticky='ew')
        self.employee_tree.configure(xscrollcommand=hscrollbar.set)
        
        # 绑定选择事件
        self.employee_tree.bind('<<TreeviewSelect>>', self.select_employee)
        
    def load_employee_data(self):
        # 清空表格
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)
        
        # 获取所有员工数据
        employees = database.fetch_employees()
        
        # 填充表格
        for emp in employees:
            # 格式化日期
            birth_date = emp[3].strftime('%Y-%m-%d') if emp[3] else ''
            hire_date = emp[5].strftime('%Y-%m-%d') if emp[5] else ''
            
            self.employee_tree.insert('', 'end', values=(emp[0], emp[1], emp[2], birth_date, emp[4], hire_date, emp[6], emp[7]))
    
    def select_employee(self, event):
        # 获取选中的行
        selected_item = self.employee_tree.selection()
        if selected_item:
            # 获取选中行的值
            row = self.employee_tree.item(selected_item[0])['values']
            
            # 清空表单
            self.clear_employee_form()
            
            # 填充表单
            self.employee_id_entry.insert(0, row[0])
            self.employee_name_entry.insert(0, row[1])
            self.employee_gender_box.set(row[2] if row[2] else self.employee_gender_box.get())
            self.employee_birth_entry.insert(0, row[3] if row[3] else '')
            self.employee_contact_entry.insert(0, row[4] if row[4] else '')
            self.employee_hire_entry.insert(0, row[5] if row[5] else '')
            
            # 选择部门和职位
            if row[6]:
                for i, dept in enumerate(self.department_names):
                    if dept == row[6]:
                        self.employee_dept_box.set(self.department_options[i])
                        break
            
            if row[7]:
                for i, pos in enumerate(self.position_names):
                    if pos == row[7]:
                        self.employee_position_box.set(self.position_options[i])
                        break
    
    def clear_employee_form(self):
        # 清空表单
        self.employee_id_entry.delete(0, 'end')
        self.employee_name_entry.delete(0, 'end')
        self.employee_birth_entry.delete(0, 'end')
        self.employee_contact_entry.delete(0, 'end')
        self.employee_hire_entry.delete(0, 'end')
        
        # 重置下拉框
        if self.department_options:
            self.employee_dept_box.set(self.department_options[0])
        if self.position_options:
            self.employee_position_box.set(self.position_options[0])
        
        # 聚焦到ID输入框
        self.employee_id_entry.focus()
    
    def add_employee(self):
        # 获取表单数据
        employee_id = self.employee_id_entry.get()
        name = self.employee_name_entry.get()
        gender = self.employee_gender_box.get()
        birth_date = self.employee_birth_entry.get() or None
        contact_info = self.employee_contact_entry.get()
        hire_date = self.employee_hire_entry.get() or None
        
        # 获取选中的部门ID和职位ID
        dept_option = self.employee_dept_box.get()
        dept_id = self.department_ids[self.department_options.index(dept_option)] if dept_option in self.department_options else None
        
        position_option = self.employee_position_box.get()
        position_id = self.position_ids[self.position_options.index(position_option)] if position_option in self.position_options else None
        
        # 验证必填字段
        if not employee_id or not name:
            messagebox.showerror('错误', '员工ID和姓名为必填项')
            return
        
        # 检查ID是否已存在
        if database.employee_exists(employee_id):
            messagebox.showerror('错误', '员工ID已存在')
            return
        
        # 日期格式验证
        try:
            if birth_date:
                birth_date = datetime.datetime.strptime(birth_date, '%Y-%m-%d').date()
            if hire_date:
                hire_date = datetime.datetime.strptime(hire_date, '%Y-%m-%d').date()
        except ValueError:
            messagebox.showerror('错误', '日期格式无效，请使用YYYY-MM-DD格式')
            return
        
        # 添加员工
        if database.insert_employee(employee_id, name, gender, birth_date, contact_info, hire_date, dept_id, position_id):
            messagebox.showinfo('成功', '员工信息已添加')
            self.clear_employee_form()
            self.load_employee_data()
    
    def update_employee(self):
        # 获取表单数据
        employee_id = self.employee_id_entry.get()
        name = self.employee_name_entry.get()
        gender = self.employee_gender_box.get()
        birth_date = self.employee_birth_entry.get() or None
        contact_info = self.employee_contact_entry.get()
        hire_date = self.employee_hire_entry.get() or None
        
        # 获取选中的部门ID和职位ID
        dept_option = self.employee_dept_box.get()
        dept_id = self.department_ids[self.department_options.index(dept_option)] if dept_option in self.department_options else None
        
        position_option = self.employee_position_box.get()
        position_id = self.position_ids[self.position_options.index(position_option)] if position_option in self.position_options else None
        
        # 验证必填字段
        if not employee_id or not name:
            messagebox.showerror('错误', '员工ID和姓名为必填项')
            return
        
        # 检查ID是否存在
        if not database.employee_exists(employee_id):
            messagebox.showerror('错误', '员工ID不存在')
            return
        
        # 日期格式验证
        try:
            if birth_date:
                birth_date = datetime.datetime.strptime(birth_date, '%Y-%m-%d').date()
            if hire_date:
                hire_date = datetime.datetime.strptime(hire_date, '%Y-%m-%d').date()
        except ValueError:
            messagebox.showerror('错误', '日期格式无效，请使用YYYY-MM-DD格式')
            return
        
        # 更新员工
        if database.update_employee(employee_id, name, gender, birth_date, contact_info, hire_date, dept_id, position_id):
            messagebox.showinfo('成功', '员工信息已更新')
            self.load_employee_data()
    
    def delete_employee(self):
        # 获取表单数据
        employee_id = self.employee_id_entry.get()
        
        # 检查ID是否存在
        if not employee_id or not database.employee_exists(employee_id):
            messagebox.showerror('错误', '请选择一个有效的员工')
            return
        
        # 确认删除
        confirm = messagebox.askyesno('确认', '确定要删除此员工吗？此操作不可撤销。')
        if confirm:
            if database.delete_employee(employee_id):
                messagebox.showinfo('成功', '员工已删除')
                self.clear_employee_form()
                self.load_employee_data()
    
    def search_employee(self):
        # 获取搜索条件
        search_by = self.employee_search_box.get()
        search_value = self.employee_search_entry.get()
        
        if search_by == '搜索条件' or not search_value:
            messagebox.showerror('错误', '请选择搜索条件并输入搜索值')
            return
        
        # 搜索员工
        results = database.search_employees_by_keyword(search_value)
        
        # 清空表格
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)
        
        # 填充表格
        for emp in results:
            # 对于搜索结果，不需要格式化日期，因为search_employees_by_keyword返回的字段不包含日期
            # 搜索结果顺序: EmployeeID, Name, Gender, ContactInfo, DeptName, PositionName
            self.employee_tree.insert('', 'end', values=(
                emp[0],  # EmployeeID
                emp[1],  # Name
                emp[2],  # Gender
                '',      # BirthDate (不存在于搜索结果中)
                emp[3],  # ContactInfo
                '',      # HireDate (不存在于搜索结果中)
                emp[4],  # DeptName
                emp[5]   # PositionName
            ))
    
    def load_departments(self):
        departments = database.fetch_departments()
        self.department_ids = []
        self.department_names = []
        self.department_options = []
        
        for dept in departments:
            self.department_ids.append(dept[0])
            self.department_names.append(dept[1])
            self.department_options.append(f"{dept[0]} - {dept[1]}")
        
        if not self.department_options:
            self.department_options = ["无部门"]
    
    def load_positions(self):
        positions = database.fetch_positions()
        self.position_ids = []
        self.position_names = []
        self.position_options = []
        
        for pos in positions:
            self.position_ids.append(pos[0])
            self.position_names.append(pos[1])
            self.position_options.append(f"{pos[0]} - {pos[1]}")
        
        if not self.position_options:
            self.position_options = ["无职位"]

    # 部门管理选项卡
    def show_department_tab(self):
        global current_tab
        
        # 清空内容框架
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        current_tab = "department"
        
        # 创建表单框架和表格框架
        self.dept_form_frame = CTkFrame(self.content_frame, fg_color='#161C30')
        self.dept_form_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
        
        self.dept_table_frame = CTkFrame(self.content_frame)
        self.dept_table_frame.grid(row=0, column=1, padx=10, pady=10, sticky='ne')
        
        # 创建部门表单
        self.create_department_form()
        
        # 创建部门表格
        self.create_department_table()
        
        # 加载部门数据
        self.load_department_data()
    
    def create_department_form(self):
        # 部门ID
        CTkLabel(self.dept_form_frame, text='部门ID', font=('arial', 16, 'bold'), text_color='white').grid(row=0, column=0, padx=20, pady=15, sticky='w')
        self.dept_id_entry = CTkEntry(self.dept_form_frame, font=('arial', 14), width=180)
        self.dept_id_entry.grid(row=0, column=1)
        
        # 部门名称
        CTkLabel(self.dept_form_frame, text='部门名称', font=('arial', 16, 'bold'), text_color='white').grid(row=1, column=0, padx=20, pady=15, sticky='w')
        self.dept_name_entry = CTkEntry(self.dept_form_frame, font=('arial', 14), width=180)
        self.dept_name_entry.grid(row=1, column=1)
        
        # 部门经理
        CTkLabel(self.dept_form_frame, text='部门经理', font=('arial', 16, 'bold'), text_color='white').grid(row=2, column=0, padx=20, pady=15, sticky='w')
        self.load_employees_for_manager()
        self.dept_manager_box = CTkComboBox(self.dept_form_frame, values=self.manager_options, font=('arial', 14), width=180, state='readonly')
        self.dept_manager_box.grid(row=2, column=1)
        if self.manager_options:
            self.dept_manager_box.set(self.manager_options[0])
        
        # 成立日期
        CTkLabel(self.dept_form_frame, text='成立日期', font=('arial', 16, 'bold'), text_color='white').grid(row=3, column=0, padx=20, pady=15, sticky='w')
        self.dept_date_entry = CTkEntry(self.dept_form_frame, font=('arial', 14), width=180, placeholder_text='YYYY-MM-DD')
        self.dept_date_entry.grid(row=3, column=1)
        
        # 部门描述
        CTkLabel(self.dept_form_frame, text='部门描述', font=('arial', 16, 'bold'), text_color='white').grid(row=4, column=0, padx=20, pady=15, sticky='w')
        self.dept_desc_entry = CTkTextbox(self.dept_form_frame, font=('arial', 14), width=400, height=100)
        self.dept_desc_entry.grid(row=4, column=1, columnspan=2)
        
        # 按钮框架
        button_frame = CTkFrame(self.dept_form_frame, fg_color='#161C30')
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        # 新建按钮
        new_btn = CTkButton(button_frame, text='新建', font=('arial', 14, 'bold'), width=100, command=self.clear_department_form)
        new_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # 添加按钮
        add_btn = CTkButton(button_frame, text='添加', font=('arial', 14, 'bold'), width=100, command=self.add_department)
        add_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # 更新按钮
        update_btn = CTkButton(button_frame, text='更新', font=('arial', 14, 'bold'), width=100, command=self.update_department)
        update_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # 删除按钮
        delete_btn = CTkButton(button_frame, text='删除', font=('arial', 14, 'bold'), width=100, command=self.delete_department)
        delete_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 查看部门员工按钮
        view_emp_btn = CTkButton(button_frame, text='查看部门员工', font=('arial', 14, 'bold'), width=150, command=self.view_department_employees)
        view_emp_btn.grid(row=0, column=4, padx=5, pady=5)
    
    def create_department_table(self):
        # 搜索框架
        search_frame = CTkFrame(self.dept_table_frame)
        search_frame.grid(row=0, column=0, pady=5, sticky='ew')
        
        search_options = ['DeptID', 'DeptName', 'ManagerID']
        self.dept_search_box = CTkComboBox(search_frame, values=search_options, state='readonly', width=120)
        self.dept_search_box.grid(row=0, column=0, padx=5, pady=5)
        self.dept_search_box.set('搜索条件')
        
        self.dept_search_entry = CTkEntry(search_frame, width=200)
        self.dept_search_entry.grid(row=0, column=1, padx=5, pady=5)
        
        search_btn = CTkButton(search_frame, text='搜索', width=80, command=self.search_department)
        search_btn.grid(row=0, column=2, padx=5, pady=5)
        
        show_all_btn = CTkButton(search_frame, text='显示全部', width=80, command=self.load_department_data)
        show_all_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 部门表格
        columns = ('部门ID', '部门名称', '部门经理', '成立日期', '部门描述')
        
        self.dept_tree = ttk.Treeview(self.dept_table_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.dept_tree.heading(col, text=col)
            if col == '部门描述':
                self.dept_tree.column(col, width=200)
            elif col == '部门名称' or col == '部门经理':
                self.dept_tree.column(col, width=150)
            else:
                self.dept_tree.column(col, width=120)
        
        self.dept_tree.grid(row=1, column=0, sticky='nsew')
        
        # 垂直滚动条
        scrollbar = ttk.Scrollbar(self.dept_table_frame, orient=VERTICAL, command=self.dept_tree.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        self.dept_tree.configure(yscrollcommand=scrollbar.set)
        
        # 水平滚动条
        hscrollbar = ttk.Scrollbar(self.dept_table_frame, orient=HORIZONTAL, command=self.dept_tree.xview)
        hscrollbar.grid(row=2, column=0, sticky='ew')
        self.dept_tree.configure(xscrollcommand=hscrollbar.set)
        
        # 绑定选择事件
        self.dept_tree.bind('<<TreeviewSelect>>', self.select_department)
    
    def load_department_data(self):
        # 清空表格
        for item in self.dept_tree.get_children():
            self.dept_tree.delete(item)
        
        # 获取所有部门数据
        departments = database.fetch_departments()
        
        # 填充表格
        for dept in departments:
            # 查找经理姓名
            manager_name = "未指定"
            if dept[2]:  # ManagerID不为空
                manager = database.get_employee_detail(dept[2])
                if manager:
                    manager_name = manager[1]  # 员工姓名
            
            # 格式化日期
            establish_date = dept[3].strftime('%Y-%m-%d') if dept[3] else ''
            
            self.dept_tree.insert('', 'end', values=(dept[0], dept[1], manager_name, establish_date, dept[4]))
    
    def load_employees_for_manager(self):
        employees = database.fetch_employees()
        self.manager_ids = []
        self.manager_names = []
        self.manager_options = ["无经理"]  # 默认选项
        
        for emp in employees:
            self.manager_ids.append(emp[0])
            self.manager_names.append(emp[1])
            self.manager_options.append(f"{emp[0]} - {emp[1]}")
    
    def select_department(self, event):
        # 获取选中的行
        selected_item = self.dept_tree.selection()
        if selected_item:
            # 获取选中行的值
            row = self.dept_tree.item(selected_item[0])['values']
            
            # 清空表单
            self.clear_department_form()
            
            # 填充表单
            self.dept_id_entry.insert(0, row[0])
            self.dept_name_entry.insert(0, row[1])
            
            # 查找经理ID并设置下拉框
            dept_details = None
            for dept in database.fetch_departments():
                if dept[0] == row[0]:
                    dept_details = dept
                    break
            
            if dept_details and dept_details[2]:  # 如果有经理ID
                for i, emp_id in enumerate(self.manager_ids):
                    if emp_id == dept_details[2]:
                        self.dept_manager_box.set(self.manager_options[i+1])  # +1是因为第一个选项是"无经理"
                        break
            else:
                self.dept_manager_box.set(self.manager_options[0])
            
            # 填充日期和描述
            self.dept_date_entry.insert(0, row[3] if row[3] else '')
            self.dept_desc_entry.delete('0.0', 'end')
            if row[4]:
                self.dept_desc_entry.insert('0.0', row[4])
    
    def clear_department_form(self):
        # 清空表单
        self.dept_id_entry.delete(0, 'end')
        self.dept_name_entry.delete(0, 'end')
        self.dept_date_entry.delete(0, 'end')
        self.dept_desc_entry.delete('0.0', 'end')
        
        # 重置下拉框
        self.dept_manager_box.set(self.manager_options[0])
        
        # 聚焦到ID输入框
        self.dept_id_entry.focus()
    
    def add_department(self):
        # 获取表单数据
        dept_id = self.dept_id_entry.get()
        dept_name = self.dept_name_entry.get()
        manager_option = self.dept_manager_box.get()
        establish_date = self.dept_date_entry.get() or None
        description = self.dept_desc_entry.get('0.0', 'end-1c')  # 获取文本框内容并删除末尾换行符
        
        # 获取经理ID
        manager_id = None
        if manager_option != "无经理":
            manager_id = manager_option.split(' - ')[0]
        
        # 验证必填字段
        if not dept_id or not dept_name:
            messagebox.showerror('错误', '部门ID和名称为必填项')
            return
        
        # 检查ID是否已存在
        if database.dept_exists(dept_id):
            messagebox.showerror('错误', '部门ID已存在')
            return
        
        # 日期格式验证
        try:
            if establish_date:
                establish_date = datetime.datetime.strptime(establish_date, '%Y-%m-%d').date()
        except ValueError:
            messagebox.showerror('错误', '日期格式无效，请使用YYYY-MM-DD格式')
            return
        
        # 添加部门
        if database.insert_department(dept_id, dept_name, manager_id, establish_date, description):
            messagebox.showinfo('成功', '部门信息已添加')
            self.clear_department_form()
            self.load_department_data()
    
    def update_department(self):
        # 获取表单数据
        dept_id = self.dept_id_entry.get()
        dept_name = self.dept_name_entry.get()
        manager_option = self.dept_manager_box.get()
        establish_date = self.dept_date_entry.get() or None
        description = self.dept_desc_entry.get('0.0', 'end-1c')
        
        # 获取经理ID
        manager_id = None
        if manager_option != "无经理":
            manager_id = manager_option.split(' - ')[0]
        
        # 验证必填字段
        if not dept_id or not dept_name:
            messagebox.showerror('错误', '部门ID和名称为必填项')
            return
        
        # 检查ID是否存在
        if not database.dept_exists(dept_id):
            messagebox.showerror('错误', '部门ID不存在')
            return
        
        # 日期格式验证
        try:
            if establish_date:
                establish_date = datetime.datetime.strptime(establish_date, '%Y-%m-%d').date()
        except ValueError:
            messagebox.showerror('错误', '日期格式无效，请使用YYYY-MM-DD格式')
            return
        
        # 更新部门
        if database.update_department(dept_id, dept_name, manager_id, establish_date, description):
            messagebox.showinfo('成功', '部门信息已更新')
            self.load_department_data()
    
    def delete_department(self):
        # 获取表单数据
        dept_id = self.dept_id_entry.get()
        
        # 检查ID是否存在
        if not dept_id or not database.dept_exists(dept_id):
            messagebox.showerror('错误', '请选择一个有效的部门')
            return
        
        # 确认删除
        confirm = messagebox.askyesno('确认', '确定要删除此部门吗？此操作不可撤销。')
        if confirm:
            if database.delete_department(dept_id):
                messagebox.showinfo('成功', '部门已删除')
                self.clear_department_form()
                self.load_department_data()
    
    def search_department(self):
        # 获取搜索条件
        search_by = self.dept_search_box.get()
        search_value = self.dept_search_entry.get()
        
        if search_by == '搜索条件' or not search_value:
            messagebox.showerror('错误', '请选择搜索条件并输入搜索值')
            return
        
        # 清空表格
        for item in self.dept_tree.get_children():
            self.dept_tree.delete(item)
        
        # 获取所有部门
        departments = database.fetch_departments()
        
        # 根据搜索条件筛选
        filtered_departments = []
        for dept in departments:
            if search_by == 'DeptID' and search_value.lower() in dept[0].lower():
                filtered_departments.append(dept)
            elif search_by == 'DeptName' and search_value.lower() in dept[1].lower():
                filtered_departments.append(dept)
            elif search_by == 'ManagerID' and dept[2] and search_value.lower() in dept[2].lower():
                filtered_departments.append(dept)
        
        # 填充表格
        for dept in filtered_departments:
            # 查找经理姓名
            manager_name = "未指定"
            if dept[2]:  # ManagerID不为空
                manager = database.get_employee_detail(dept[2])
                if manager:
                    manager_name = manager[1]  # 员工姓名
            
            # 格式化日期
            establish_date = dept[3].strftime('%Y-%m-%d') if dept[3] else ''
            
            self.dept_tree.insert('', 'end', values=(dept[0], dept[1], manager_name, establish_date, dept[4]))
    
    def view_department_employees(self):
        # 获取选中的部门ID
        dept_id = self.dept_id_entry.get()
        
        if not dept_id or not database.dept_exists(dept_id):
            messagebox.showerror('错误', '请选择一个有效的部门')
            return
        
        # 创建新窗口显示部门员工
        emp_window = CTkToplevel(self.root)
        emp_window.title(f"部门员工 - {self.dept_name_entry.get()}")
        emp_window.geometry('800x500')
        emp_window.grab_set()  # 使窗口模态
        
        # 创建表格
        columns = ('员工ID', '姓名', '性别', '职位')
        emp_tree = ttk.Treeview(emp_window, columns=columns, show='headings', height=15)
        for col in columns:
            emp_tree.heading(col, text=col)
            emp_tree.column(col, width=180)
        
        emp_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(emp_window, orient='vertical', command=emp_tree.yview)
        scrollbar.place(relx=1, rely=0, relheight=1, anchor='ne')
        emp_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取并显示部门员工
        employees = database.get_department_employees(dept_id)
        for emp in employees:
            emp_tree.insert('', 'end', values=(emp[0], emp[1], emp[2], emp[3]))
        
        # 添加关闭按钮
        close_btn = CTkButton(emp_window, text='关闭', command=emp_window.destroy)
        close_btn.pack(pady=10)

    # 职位管理选项卡
    def show_position_tab(self):
        global current_tab
        
        # 清空内容框架
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        current_tab = "position"
        
        # 创建表单框架和表格框架
        self.pos_form_frame = CTkFrame(self.content_frame, fg_color='#161C30')
        self.pos_form_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
        
        self.pos_table_frame = CTkFrame(self.content_frame)
        self.pos_table_frame.grid(row=0, column=1, padx=10, pady=10, sticky='ne')
        
        # 创建职位表单
        self.create_position_form()
        
        # 创建职位表格
        self.create_position_table()
        
        # 加载职位数据
        self.load_position_data()
    
    def create_position_form(self):
        # 职位ID
        CTkLabel(self.pos_form_frame, text='职位ID', font=('arial', 16, 'bold'), text_color='white').grid(row=0, column=0, padx=20, pady=15, sticky='w')
        self.pos_id_entry = CTkEntry(self.pos_form_frame, font=('arial', 14), width=180)
        self.pos_id_entry.grid(row=0, column=1)
        
        # 职位名称
        CTkLabel(self.pos_form_frame, text='职位名称', font=('arial', 16, 'bold'), text_color='white').grid(row=1, column=0, padx=20, pady=15, sticky='w')
        self.pos_name_entry = CTkEntry(self.pos_form_frame, font=('arial', 14), width=180)
        self.pos_name_entry.grid(row=1, column=1)
        
        # 职位类别
        CTkLabel(self.pos_form_frame, text='职位类别', font=('arial', 16, 'bold'), text_color='white').grid(row=2, column=0, padx=20, pady=15, sticky='w')
        category_options = ['技术', '管理', '市场', '行政', '财务', '人力资源', '其他']
        self.pos_category_box = CTkComboBox(self.pos_form_frame, values=category_options, font=('arial', 14), width=180, state='readonly')
        self.pos_category_box.grid(row=2, column=1)
        self.pos_category_box.set(category_options[0])
        
        # 职位级别
        CTkLabel(self.pos_form_frame, text='职位级别', font=('arial', 16, 'bold'), text_color='white').grid(row=3, column=0, padx=20, pady=15, sticky='w')
        level_options = ['1', '2', '3', '4', '5']
        self.pos_level_box = CTkComboBox(self.pos_form_frame, values=level_options, font=('arial', 14), width=180, state='readonly')
        self.pos_level_box.grid(row=3, column=1)
        self.pos_level_box.set(level_options[0])
        
        # 基本工资
        CTkLabel(self.pos_form_frame, text='基本工资', font=('arial', 16, 'bold'), text_color='white').grid(row=4, column=0, padx=20, pady=15, sticky='w')
        self.pos_salary_entry = CTkEntry(self.pos_form_frame, font=('arial', 14), width=180)
        self.pos_salary_entry.grid(row=4, column=1)
        
        # 按钮框架
        button_frame = CTkFrame(self.pos_form_frame, fg_color='#161C30')
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        # 新建按钮
        new_btn = CTkButton(button_frame, text='新建', font=('arial', 14, 'bold'), width=100, command=self.clear_position_form)
        new_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # 添加按钮
        add_btn = CTkButton(button_frame, text='添加', font=('arial', 14, 'bold'), width=100, command=self.add_position)
        add_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # 更新按钮
        update_btn = CTkButton(button_frame, text='更新', font=('arial', 14, 'bold'), width=100, command=self.update_position)
        update_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # 删除按钮
        delete_btn = CTkButton(button_frame, text='删除', font=('arial', 14, 'bold'), width=100, command=self.delete_position)
        delete_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 查看职位员工按钮
        view_emp_btn = CTkButton(button_frame, text='查看职位员工', font=('arial', 14, 'bold'), width=150, command=self.view_position_employees)
        view_emp_btn.grid(row=0, column=4, padx=5, pady=5)
    
    def create_position_table(self):
        # 搜索框架
        search_frame = CTkFrame(self.pos_table_frame)
        search_frame.grid(row=0, column=0, pady=5, sticky='ew')
        
        search_options = ['PositionID', 'PositionName', 'PositionCategory']
        self.pos_search_box = CTkComboBox(search_frame, values=search_options, state='readonly', width=120)
        self.pos_search_box.grid(row=0, column=0, padx=5, pady=5)
        self.pos_search_box.set('搜索条件')
        
        self.pos_search_entry = CTkEntry(search_frame, width=200)
        self.pos_search_entry.grid(row=0, column=1, padx=5, pady=5)
        
        search_btn = CTkButton(search_frame, text='搜索', width=80, command=self.search_position)
        search_btn.grid(row=0, column=2, padx=5, pady=5)
        
        show_all_btn = CTkButton(search_frame, text='显示全部', width=80, command=self.load_position_data)
        show_all_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 职位表格
        columns = ('职位ID', '职位名称', '职位类别', '职位级别', '基本工资')
        
        self.pos_tree = ttk.Treeview(self.pos_table_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.pos_tree.heading(col, text=col)
            if col == '职位名称':
                self.pos_tree.column(col, width=180)
            else:
                self.pos_tree.column(col, width=120)
        
        self.pos_tree.grid(row=1, column=0, sticky='nsew')
        
        # A垂直滚动条
        scrollbar = ttk.Scrollbar(self.pos_table_frame, orient=VERTICAL, command=self.pos_tree.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        self.pos_tree.configure(yscrollcommand=scrollbar.set)
        
        # 水平滚动条
        hscrollbar = ttk.Scrollbar(self.pos_table_frame, orient=HORIZONTAL, command=self.pos_tree.xview)
        hscrollbar.grid(row=2, column=0, sticky='ew')
        self.pos_tree.configure(xscrollcommand=hscrollbar.set)
        
        # 绑定选择事件
        self.pos_tree.bind('<<TreeviewSelect>>', self.select_position)
    
    def load_position_data(self):
        # 清空表格
        for item in self.pos_tree.get_children():
            self.pos_tree.delete(item)
        
        # 获取所有职位数据
        positions = database.fetch_positions()
        
        # 填充表格
        for pos in positions:
            self.pos_tree.insert('', 'end', values=(pos[0], pos[1], pos[2], pos[3], pos[4]))
    
    def select_position(self, event):
        # 获取选中的行
        selected_item = self.pos_tree.selection()
        if selected_item:
            # 获取选中行的值
            row = self.pos_tree.item(selected_item[0])['values']
            
            # 清空表单
            self.clear_position_form()
            
            # 填充表单
            self.pos_id_entry.insert(0, row[0])
            self.pos_name_entry.insert(0, row[1])
            
            # 设置类别和级别
            if row[2]:
                self.pos_category_box.set(row[2])
            if row[3]:
                self.pos_level_box.set(str(row[3]))
            
            # 设置基本工资
            if row[4]:
                self.pos_salary_entry.insert(0, str(row[4]))
    
    def clear_position_form(self):
        # 清空表单
        self.pos_id_entry.delete(0, 'end')
        self.pos_name_entry.delete(0, 'end')
        self.pos_salary_entry.delete(0, 'end')
        
        # 重置下拉框
        self.pos_category_box.set('技术')
        self.pos_level_box.set('1')
        
        # 聚焦到ID输入框
        self.pos_id_entry.focus()
    
    def add_position(self):
        # 获取表单数据
        position_id = self.pos_id_entry.get()
        position_name = self.pos_name_entry.get()
        category = self.pos_category_box.get()
        level = self.pos_level_box.get()
        base_salary = self.pos_salary_entry.get() or None
        
        # 验证必填字段
        if not position_id or not position_name:
            messagebox.showerror('错误', '职位ID和名称为必填项')
            return
        
        # 检查ID是否已存在
        if database.position_exists(position_id):
            messagebox.showerror('错误', '职位ID已存在')
            return
        
        # 验证工资
        try:
            if base_salary:
                base_salary = float(base_salary)
        except ValueError:
            messagebox.showerror('错误', '基本工资必须是数字')
            return
        
        # 添加职位
        if database.insert_position(position_id, position_name, category, int(level), base_salary):
            messagebox.showinfo('成功', '职位信息已添加')
            self.clear_position_form()
            self.load_position_data()
    
    def update_position(self):
        # 获取表单数据
        position_id = self.pos_id_entry.get()
        position_name = self.pos_name_entry.get()
        category = self.pos_category_box.get()
        level = self.pos_level_box.get()
        base_salary = self.pos_salary_entry.get() or None
        
        # 验证必填字段
        if not position_id or not position_name:
            messagebox.showerror('错误', '职位ID和名称为必填项')
            return
        
        # 检查ID是否存在
        if not database.position_exists(position_id):
            messagebox.showerror('错误', '职位ID不存在')
            return
        
        # 验证工资
        try:
            if base_salary:
                base_salary = float(base_salary)
        except ValueError:
            messagebox.showerror('错误', '基本工资必须是数字')
            return
        
        # 更新职位
        if database.update_position(position_id, position_name, category, int(level), base_salary):
            messagebox.showinfo('成功', '职位信息已更新')
            self.load_position_data()
    
    def delete_position(self):
        # 获取表单数据
        position_id = self.pos_id_entry.get()
        
        # 检查ID是否存在
        if not position_id or not database.position_exists(position_id):
            messagebox.showerror('错误', '请选择一个有效的职位')
            return
        
        # 确认删除
        confirm = messagebox.askyesno('确认', '确定要删除此职位吗？此操作不可撤销。')
        if confirm:
            if database.delete_position(position_id):
                messagebox.showinfo('成功', '职位已删除')
                self.clear_position_form()
                self.load_position_data()
    
    def search_position(self):
        # 获取搜索条件
        search_by = self.pos_search_box.get()
        search_value = self.pos_search_entry.get()
        
        if search_by == '搜索条件' or not search_value:
            messagebox.showerror('错误', '请选择搜索条件并输入搜索值')
            return
        
        # 清空表格
        for item in self.pos_tree.get_children():
            self.pos_tree.delete(item)
        
        # 获取所有职位
        positions = database.fetch_positions()
        
        # 根据搜索条件筛选
        filtered_positions = []
        for pos in positions:
            if search_by == 'PositionID' and search_value.lower() in str(pos[0]).lower():
                filtered_positions.append(pos)
            elif search_by == 'PositionName' and search_value.lower() in str(pos[1]).lower():
                filtered_positions.append(pos)
            elif search_by == 'PositionCategory' and pos[2] and search_value.lower() in str(pos[2]).lower():
                filtered_positions.append(pos)
        
        # 填充表格
        for pos in filtered_positions:
            self.pos_tree.insert('', 'end', values=(pos[0], pos[1], pos[2], pos[3], pos[4]))
    
    def view_position_employees(self):
        # 获取选中的职位ID
        position_id = self.pos_id_entry.get()
        
        if not position_id or not database.position_exists(position_id):
            messagebox.showerror('错误', '请选择一个有效的职位')
            return
        
        # 创建新窗口显示职位员工
        emp_window = CTkToplevel(self.root)
        emp_window.title(f"职位员工 - {self.pos_name_entry.get()}")
        emp_window.geometry('800x500')
        emp_window.grab_set()  # 使窗口模态
        
        # 创建表格
        columns = ('员工ID', '姓名', '性别', '部门')
        emp_tree = ttk.Treeview(emp_window, columns=columns, show='headings', height=15)
        for col in columns:
            emp_tree.heading(col, text=col)
            emp_tree.column(col, width=180)
        
        emp_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(emp_window, orient='vertical', command=emp_tree.yview)
        scrollbar.place(relx=1, rely=0, relheight=1, anchor='ne')
        emp_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取并显示职位员工
        employees = database.get_position_employees(position_id)
        for emp in employees:
            emp_tree.insert('', 'end', values=(emp[0], emp[1], emp[2], emp[3]))
        
        # 添加关闭按钮
        close_btn = CTkButton(emp_window, text='关闭', command=emp_window.destroy)
        close_btn.pack(pady=10)

    # 项目管理选项卡
    def show_project_tab(self):
        global current_tab
        
        # 清空内容框架
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        current_tab = "project"
        
        # 创建表单框架和表格框架
        self.proj_form_frame = CTkFrame(self.content_frame, fg_color='#161C30')
        self.proj_form_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
        
        self.proj_table_frame = CTkFrame(self.content_frame)
        self.proj_table_frame.grid(row=0, column=1, padx=10, pady=10, sticky='ne')
        
        # 创建项目表单
        self.create_project_form()
        
        # 创建项目表格
        self.create_project_table()
        
        # 加载项目数据
        self.load_project_data()
    
    def create_project_form(self):
        # 项目ID
        CTkLabel(self.proj_form_frame, text='项目ID', font=('arial', 16, 'bold'), text_color='white').grid(row=0, column=0, padx=20, pady=15, sticky='w')
        self.proj_id_entry = CTkEntry(self.proj_form_frame, font=('arial', 14), width=180)
        self.proj_id_entry.grid(row=0, column=1)
        
        # 项目名称
        CTkLabel(self.proj_form_frame, text='项目名称', font=('arial', 16, 'bold'), text_color='white').grid(row=1, column=0, padx=20, pady=15, sticky='w')
        self.proj_name_entry = CTkEntry(self.proj_form_frame, font=('arial', 14), width=180)
        self.proj_name_entry.grid(row=1, column=1)
        
        # 开始日期
        CTkLabel(self.proj_form_frame, text='开始日期', font=('arial', 16, 'bold'), text_color='white').grid(row=2, column=0, padx=20, pady=15, sticky='w')
        self.proj_start_entry = CTkEntry(self.proj_form_frame, font=('arial', 14), width=180, placeholder_text='YYYY-MM-DD')
        self.proj_start_entry.grid(row=2, column=1)
        
        # 结束日期
        CTkLabel(self.proj_form_frame, text='结束日期', font=('arial', 16, 'bold'), text_color='white').grid(row=3, column=0, padx=20, pady=15, sticky='w')
        self.proj_end_entry = CTkEntry(self.proj_form_frame, font=('arial', 14), width=180, placeholder_text='YYYY-MM-DD')
        self.proj_end_entry.grid(row=3, column=1)
        
        # 项目状态
        CTkLabel(self.proj_form_frame, text='项目状态', font=('arial', 16, 'bold'), text_color='white').grid(row=4, column=0, padx=20, pady=15, sticky='w')
        status_options = ['未开始', '进行中', '已完成', '已暂停', '已取消']
        self.proj_status_box = CTkComboBox(self.proj_form_frame, values=status_options, font=('arial', 14), width=180, state='readonly')
        self.proj_status_box.grid(row=4, column=1)
        self.proj_status_box.set(status_options[0])
        
        # 项目负责人
        CTkLabel(self.proj_form_frame, text='项目负责人', font=('arial', 16, 'bold'), text_color='white').grid(row=5, column=0, padx=20, pady=15, sticky='w')
        self.load_employees_for_leader()
        self.proj_leader_box = CTkComboBox(self.proj_form_frame, values=self.leader_options, font=('arial', 14), width=180, state='readonly')
        self.proj_leader_box.grid(row=5, column=1)
        if self.leader_options:
            self.proj_leader_box.set(self.leader_options[0])
        
        # 按钮框架
        button_frame = CTkFrame(self.proj_form_frame, fg_color='#161C30')
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)
        
        # 新建按钮
        new_btn = CTkButton(button_frame, text='新建', font=('arial', 14, 'bold'), width=100, command=self.clear_project_form)
        new_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # 添加按钮
        add_btn = CTkButton(button_frame, text='添加', font=('arial', 14, 'bold'), width=100, command=self.add_project)
        add_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # 更新按钮
        update_btn = CTkButton(button_frame, text='更新', font=('arial', 14, 'bold'), width=100, command=self.update_project)
        update_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # 删除按钮
        delete_btn = CTkButton(button_frame, text='删除', font=('arial', 14, 'bold'), width=100, command=self.delete_project)
        delete_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 查看项目成员按钮
        view_members_btn = CTkButton(button_frame, text='查看项目成员', font=('arial', 14, 'bold'), width=150, command=self.view_project_members)
        view_members_btn.grid(row=0, column=4, padx=5, pady=5)
    
    def create_project_table(self):
        # 搜索框架
        search_frame = CTkFrame(self.proj_table_frame)
        search_frame.grid(row=0, column=0, pady=5, sticky='ew')
        
        search_options = ['ProjectID', 'ProjectName', 'LeaderID', 'Status']
        self.proj_search_box = CTkComboBox(search_frame, values=search_options, state='readonly', width=120)
        self.proj_search_box.grid(row=0, column=0, padx=5, pady=5)
        self.proj_search_box.set('搜索条件')
        
        self.proj_search_entry = CTkEntry(search_frame, width=200)
        self.proj_search_entry.grid(row=0, column=1, padx=5, pady=5)
        
        search_btn = CTkButton(search_frame, text='搜索', width=80, command=self.search_project)
        search_btn.grid(row=0, column=2, padx=5, pady=5)
        
        show_all_btn = CTkButton(search_frame, text='显示全部', width=80, command=self.load_project_data)
        show_all_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 项目表格
        columns = ('项目ID', '项目名称', '开始日期', '结束日期', '状态', '项目负责人')
        
        self.proj_tree = ttk.Treeview(self.proj_table_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.proj_tree.heading(col, text=col)
            if col == '项目名称':
                self.proj_tree.column(col, width=200)
            elif col == '项目负责人':
                self.proj_tree.column(col, width=150)
            else:
                self.proj_tree.column(col, width=120)
        
        self.proj_tree.grid(row=1, column=0, sticky='nsew')
        
        # 垂直滚动条
        scrollbar = ttk.Scrollbar(self.proj_table_frame, orient=VERTICAL, command=self.proj_tree.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        self.proj_tree.configure(yscrollcommand=scrollbar.set)
        
        # 水平滚动条
        hscrollbar = ttk.Scrollbar(self.proj_table_frame, orient=HORIZONTAL, command=self.proj_tree.xview)
        hscrollbar.grid(row=2, column=0, sticky='ew')
        self.proj_tree.configure(xscrollcommand=hscrollbar.set)
        
        # 绑定选择事件
        self.proj_tree.bind('<<TreeviewSelect>>', self.select_project)
        
        # 添加查看项目成员按钮
        view_btn = CTkButton(self.proj_table_frame, text='查看项目成员', command=self.view_project_members)
        view_btn.grid(row=3, column=0, pady=10)
    
    def load_project_data(self):
        # 清空表格
        for item in self.proj_tree.get_children():
            self.proj_tree.delete(item)
        
        # 获取所有项目数据
        projects = database.fetch_projects()
        
        # 填充表格
        for proj in projects:
            # 格式化日期
            start_date = proj[2].strftime('%Y-%m-%d') if proj[2] else ''
            end_date = proj[3].strftime('%Y-%m-%d') if proj[3] else ''
            
            self.proj_tree.insert('', 'end', values=(proj[0], proj[1], start_date, end_date, proj[4], proj[5]))
    
    def load_employees_for_leader(self):
        employees = database.fetch_employees()
        self.leader_ids = []
        self.leader_names = []
        self.leader_options = ["无负责人"]  # 默认选项
        
        for emp in employees:
            self.leader_ids.append(emp[0])
            self.leader_names.append(emp[1])
            self.leader_options.append(f"{emp[0]} - {emp[1]}")
    
    def select_project(self, event):
        # 获取选中的行
        selected_item = self.proj_tree.selection()
        if selected_item:
            # 获取选中行的值
            row = self.proj_tree.item(selected_item[0])['values']
            
            # 清空表单
            self.clear_project_form()
            
            # 填充表单
            self.proj_id_entry.insert(0, row[0])
            self.proj_name_entry.insert(0, row[1])
            self.proj_start_entry.insert(0, row[2] if row[2] else '')
            self.proj_end_entry.insert(0, row[3] if row[3] else '')
            
            # 设置状态
            if row[4]:
                self.proj_status_box.set(row[4])
            
            # 查找负责人ID并设置下拉框
            proj_details = None
            for proj in database.fetch_projects():
                if proj[0] == row[0]:
                    proj_details = proj
                    break
            
            # 设置负责人
            leader_name = row[5]
            if leader_name:
                for i, option in enumerate(self.leader_options[1:], 1):  # 跳过"无负责人"选项
                    if leader_name in option:
                        self.proj_leader_box.set(option)
                        break
            else:
                self.proj_leader_box.set(self.leader_options[0])
    
    def clear_project_form(self):
        # 清空表单
        self.proj_id_entry.delete(0, 'end')
        self.proj_name_entry.delete(0, 'end')
        self.proj_start_entry.delete(0, 'end')
        self.proj_end_entry.delete(0, 'end')
        
        # 重置下拉框
        self.proj_status_box.set('未开始')
        self.proj_leader_box.set(self.leader_options[0])
        
        # 聚焦到ID输入框
        self.proj_id_entry.focus()
    
    def add_project(self):
        # 获取表单数据
        project_id = self.proj_id_entry.get()
        project_name = self.proj_name_entry.get()
        start_date = self.proj_start_entry.get() or None
        end_date = self.proj_end_entry.get() or None
        status = self.proj_status_box.get()
        
        # 获取负责人ID
        leader_option = self.proj_leader_box.get()
        leader_id = None
        if leader_option != "无负责人":
            leader_id = leader_option.split(' - ')[0]
        
        # 验证必填字段
        if not project_id or not project_name:
            messagebox.showerror('错误', '项目ID和名称为必填项')
            return
        
        # 检查ID是否已存在
        if database.project_exists(project_id):
            messagebox.showerror('错误', '项目ID已存在')
            return
        
        # 日期格式验证
        try:
            if start_date:
                start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d').date()
            if end_date:
                end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d').date()
                
                # 验证结束日期是否晚于开始日期
                if start_date and end_date < start_date:
                    messagebox.showerror('错误', '结束日期不能早于开始日期')
                    return
        except ValueError:
            messagebox.showerror('错误', '日期格式无效，请使用YYYY-MM-DD格式')
            return
        
        # 添加项目
        if database.insert_project(project_id, project_name, start_date, end_date, status, leader_id):
            messagebox.showinfo('成功', '项目信息已添加')
            self.clear_project_form()
            self.load_project_data()
    
    def update_project(self):
        # 获取表单数据
        project_id = self.proj_id_entry.get()
        project_name = self.proj_name_entry.get()
        start_date = self.proj_start_entry.get() or None
        end_date = self.proj_end_entry.get() or None
        status = self.proj_status_box.get()
        
        # 获取负责人ID
        leader_option = self.proj_leader_box.get()
        leader_id = None
        if leader_option != "无负责人":
            leader_id = leader_option.split(' - ')[0]
        
        # 验证必填字段
        if not project_id or not project_name:
            messagebox.showerror('错误', '项目ID和名称为必填项')
            return
        
        # 检查ID是否存在
        if not database.project_exists(project_id):
            messagebox.showerror('错误', '项目ID不存在')
            return
        
        # 日期格式验证
        try:
            if start_date:
                start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d').date()
            if end_date:
                end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d').date()
                
                # 验证结束日期是否晚于开始日期
                if start_date and end_date < start_date:
                    messagebox.showerror('错误', '结束日期不能早于开始日期')
                    return
        except ValueError:
            messagebox.showerror('错误', '日期格式无效，请使用YYYY-MM-DD格式')
            return
        
        # 更新项目
        if database.update_project(project_id, project_name, start_date, end_date, status, leader_id):
            messagebox.showinfo('成功', '项目信息已更新')
            self.load_project_data()
    
    def delete_project(self):
        # 获取表单数据
        project_id = self.proj_id_entry.get()
        
        # 检查ID是否存在
        if not project_id or not database.project_exists(project_id):
            messagebox.showerror('错误', '请选择一个有效的项目')
            return
        
        # 确认删除
        confirm = messagebox.askyesno('确认', '确定要删除此项目吗？此操作不可撤销，并将删除与此项目相关的所有参与记录。')
        if confirm:
            if database.delete_project(project_id):
                messagebox.showinfo('成功', '项目已删除')
                self.clear_project_form()
                self.load_project_data()
    
    def search_project(self):
        # 获取搜索条件
        search_by = self.proj_search_box.get()
        search_value = self.proj_search_entry.get()
        
        if search_by == '搜索条件' or not search_value:
            messagebox.showerror('错误', '请选择搜索条件并输入搜索值')
            return
        
        # 清空表格
        for item in self.proj_tree.get_children():
            self.proj_tree.delete(item)
        
        # 获取所有项目
        projects = database.fetch_projects()
        
        # 根据搜索条件筛选
        filtered_projects = []
        for proj in projects:
            if search_by == 'ProjectID' and search_value.lower() in str(proj[0]).lower():
                filtered_projects.append(proj)
            elif search_by == 'ProjectName' and search_value.lower() in str(proj[1]).lower():
                filtered_projects.append(proj)
            elif search_by == 'Status' and proj[4] and search_value.lower() in str(proj[4]).lower():
                filtered_projects.append(proj)
            elif search_by == 'Leader' and proj[5] and search_value.lower() in str(proj[5]).lower():
                filtered_projects.append(proj)
        
        # 填充表格
        for proj in filtered_projects:
            # 格式化日期
            start_date = proj[2].strftime('%Y-%m-%d') if proj[2] else ''
            end_date = proj[3].strftime('%Y-%m-%d') if proj[3] else ''
            
            self.proj_tree.insert('', 'end', values=(proj[0], proj[1], start_date, end_date, proj[4], proj[5]))
    
    def view_project_members(self):
        # 获取选中的项目ID
        project_id = self.proj_id_entry.get()
        
        if not project_id or not database.project_exists(project_id):
            messagebox.showerror('错误', '请选择一个有效的项目')
            return
        
        # 创建新窗口显示项目成员
        members_window = CTkToplevel(self.root)
        members_window.title(f"项目成员 - {self.proj_name_entry.get()}")
        members_window.geometry('800x500')
        members_window.grab_set()  # 使窗口模态
        
        # 创建表格
        columns = ('员工ID', '姓名', '项目角色', '开始日期', '结束日期')
        members_tree = ttk.Treeview(members_window, columns=columns, show='headings', height=15)
        for col in columns:
            members_tree.heading(col, text=col)
            members_tree.column(col, width=150)
        
        members_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(members_window, orient='vertical', command=members_tree.yview)
        scrollbar.place(relx=1, rely=0, relheight=1, anchor='ne')
        members_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取并显示项目成员
        members = database.fetch_project_members(project_id)
        for member in members:
            # 格式化日期
            start_date = member[3].strftime('%Y-%m-%d') if member[3] else ''
            end_date = member[4].strftime('%Y-%m-%d') if member[4] else ''
            
            members_tree.insert('', 'end', values=(member[0], member[1], member[2], start_date, end_date))
        
        # 添加关闭按钮
        close_btn = CTkButton(members_window, text='关闭', command=members_window.destroy)
        close_btn.pack(pady=10)

    # 项目参与选项卡
    def show_participation_tab(self):
        global current_tab
        
        # 清空内容框架
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        current_tab = "participation"
        
        # 创建表单框架和表格框架
        self.part_form_frame = CTkFrame(self.content_frame, fg_color='#161C30')
        self.part_form_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
        
        self.part_table_frame = CTkFrame(self.content_frame)
        self.part_table_frame.grid(row=0, column=1, padx=10, pady=10, sticky='ne')
        
        # 创建参与表单
        self.create_participation_form()
        
        # 创建参与表格
        self.create_participation_table()
        
        # 加载参与数据
        self.load_participation_data()
    
    def create_participation_form(self):
        # 员工
        CTkLabel(self.part_form_frame, text='员工', font=('arial', 16, 'bold'), text_color='white').grid(row=0, column=0, padx=20, pady=15, sticky='w')
        self.load_employees_for_participation()
        self.part_employee_box = CTkComboBox(self.part_form_frame, values=self.part_employee_options, font=('arial', 14), width=220, state='readonly')
        self.part_employee_box.grid(row=0, column=1)
        if self.part_employee_options:
            self.part_employee_box.set(self.part_employee_options[0])
        
        # 项目
        CTkLabel(self.part_form_frame, text='项目', font=('arial', 16, 'bold'), text_color='white').grid(row=1, column=0, padx=20, pady=15, sticky='w')
        self.load_projects_for_participation()
        self.part_project_box = CTkComboBox(self.part_form_frame, values=self.part_project_options, font=('arial', 14), width=220, state='readonly')
        self.part_project_box.grid(row=1, column=1)
        if self.part_project_options:
            self.part_project_box.set(self.part_project_options[0])
        
        # 项目角色
        CTkLabel(self.part_form_frame, text='项目角色', font=('arial', 16, 'bold'), text_color='white').grid(row=2, column=0, padx=20, pady=15, sticky='w')
        role_options = ['项目经理', '技术负责人', '开发人员', '测试人员', '设计人员', '产品经理', '其他']
        self.part_role_box = CTkComboBox(self.part_form_frame, values=role_options, font=('arial', 14), width=220, state='readonly')
        self.part_role_box.grid(row=2, column=1)
        self.part_role_box.set(role_options[0])
        
        # 开始日期
        CTkLabel(self.part_form_frame, text='开始日期', font=('arial', 16, 'bold'), text_color='white').grid(row=3, column=0, padx=20, pady=15, sticky='w')
        self.part_start_entry = CTkEntry(self.part_form_frame, font=('arial', 14), width=220, placeholder_text='YYYY-MM-DD')
        self.part_start_entry.grid(row=3, column=1)
        
        # 结束日期
        CTkLabel(self.part_form_frame, text='结束日期', font=('arial', 16, 'bold'), text_color='white').grid(row=4, column=0, padx=20, pady=15, sticky='w')
        self.part_end_entry = CTkEntry(self.part_form_frame, font=('arial', 14), width=220, placeholder_text='YYYY-MM-DD (可选)')
        self.part_end_entry.grid(row=4, column=1)
        
        # 按钮框架
        button_frame = CTkFrame(self.part_form_frame, fg_color='#161C30')
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        # 新建按钮
        new_btn = CTkButton(button_frame, text='新建', font=('arial', 14, 'bold'), width=100, command=self.clear_participation_form)
        new_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # 添加按钮
        add_btn = CTkButton(button_frame, text='添加', font=('arial', 14, 'bold'), width=100, command=self.add_participation)
        add_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # 更新按钮
        update_btn = CTkButton(button_frame, text='更新', font=('arial', 14, 'bold'), width=100, command=self.update_participation)
        update_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # 删除按钮
        delete_btn = CTkButton(button_frame, text='删除', font=('arial', 14, 'bold'), width=100, command=self.delete_participation)
        delete_btn.grid(row=0, column=3, padx=5, pady=5)
    
    def create_participation_table(self):
        # 搜索框架
        search_frame = CTkFrame(self.part_table_frame)
        search_frame.grid(row=0, column=0, pady=5, sticky='ew')
        
        search_options = ['Employee', 'Project', 'Role']
        self.part_search_box = CTkComboBox(search_frame, values=search_options, state='readonly', width=120)
        self.part_search_box.grid(row=0, column=0, padx=5, pady=5)
        self.part_search_box.set('搜索条件')
        
        self.part_search_entry = CTkEntry(search_frame, width=200)
        self.part_search_entry.grid(row=0, column=1, padx=5, pady=5)
        
        search_btn = CTkButton(search_frame, text='搜索', width=80, command=self.search_participation)
        search_btn.grid(row=0, column=2, padx=5, pady=5)
        
        show_all_btn = CTkButton(search_frame, text='显示全部', width=80, command=self.load_participation_data)
        show_all_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 参与表格
        columns = ('员工ID', '员工姓名', '项目ID', '项目名称', '项目角色', '开始日期', '结束日期')
        
        self.part_tree = ttk.Treeview(self.part_table_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.part_tree.heading(col, text=col)
            if col in ['项目名称', '项目角色']:
                self.part_tree.column(col, width=180)
            elif col == '员工姓名':
                self.part_tree.column(col, width=150)
            else:
                self.part_tree.column(col, width=120)
        
        self.part_tree.grid(row=1, column=0, sticky='nsew')
        
        # 垂直滚动条
        scrollbar = ttk.Scrollbar(self.part_table_frame, orient=VERTICAL, command=self.part_tree.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        self.part_tree.configure(yscrollcommand=scrollbar.set)
        
        # 水平滚动条
        hscrollbar = ttk.Scrollbar(self.part_table_frame, orient=HORIZONTAL, command=self.part_tree.xview)
        hscrollbar.grid(row=2, column=0, sticky='ew')
        self.part_tree.configure(xscrollcommand=hscrollbar.set)
        
        # 绑定选择事件
        self.part_tree.bind('<<TreeviewSelect>>', self.select_participation)
    
    def load_participation_data(self):
        # 清空表格
        for item in self.part_tree.get_children():
            self.part_tree.delete(item)
        
        # 获取所有参与数据
        participations = database.fetch_participations()
        
        # 填充表格
        for part in participations:
            # 格式化日期
            start_date = part[5].strftime('%Y-%m-%d') if part[5] else ''
            end_date = part[6].strftime('%Y-%m-%d') if part[6] else ''
            
            self.part_tree.insert('', 'end', values=(part[0], part[1], part[2], part[3], part[4], start_date, end_date))
    
    def load_employees_for_participation(self):
        employees = database.fetch_employees()
        self.part_employee_ids = []
        self.part_employee_names = []
        self.part_employee_options = []
        
        for emp in employees:
            self.part_employee_ids.append(emp[0])
            self.part_employee_names.append(emp[1])
            self.part_employee_options.append(f"{emp[0]} - {emp[1]}")
        
        if not self.part_employee_options:
            self.part_employee_options = ["无员工"]
    
    def load_projects_for_participation(self):
        projects = database.fetch_projects()
        self.part_project_ids = []
        self.part_project_names = []
        self.part_project_options = []
        
        for proj in projects:
            self.part_project_ids.append(proj[0])
            self.part_project_names.append(proj[1])
            self.part_project_options.append(f"{proj[0]} - {proj[1]}")
        
        if not self.part_project_options:
            self.part_project_options = ["无项目"]
    
    def select_participation(self, event):
        # 获取选中的行
        selected_item = self.part_tree.selection()
        if selected_item:
            # 获取选中行的值
            row = self.part_tree.item(selected_item[0])['values']
            
            # 清空表单
            self.clear_participation_form()
            
            # 设置员工和项目下拉框
            emp_id = row[0]
            proj_id = row[2]
            
            # 设置员工
            for i, emp_id_option in enumerate(self.part_employee_ids):
                if emp_id_option == emp_id:
                    self.part_employee_box.set(self.part_employee_options[i])
                    break
            
            # 设置项目
            for i, proj_id_option in enumerate(self.part_project_ids):
                if proj_id_option == proj_id:
                    self.part_project_box.set(self.part_project_options[i])
                    break
            
            # 设置角色
            self.part_role_box.set(row[4] if row[4] else '开发人员')
            
            # 填充日期
            self.part_start_entry.insert(0, row[5] if row[5] else '')
            self.part_end_entry.insert(0, row[6] if row[6] else '')
    
    def clear_participation_form(self):
        # 清空日期输入框
        self.part_start_entry.delete(0, 'end')
        self.part_end_entry.delete(0, 'end')
        
        # 重置下拉框
        if self.part_employee_options:
            self.part_employee_box.set(self.part_employee_options[0])
        if self.part_project_options:
            self.part_project_box.set(self.part_project_options[0])
        self.part_role_box.set('开发人员')
    
    def add_participation(self):
        # 获取表单数据
        employee_option = self.part_employee_box.get()
        project_option = self.part_project_box.get()
        role = self.part_role_box.get()
        start_date = self.part_start_entry.get() or None
        end_date = self.part_end_entry.get() or None
        
        # 获取员工ID和项目ID
        if employee_option == "无员工" or project_option == "无项目":
            messagebox.showerror('错误', '请选择有效的员工和项目')
            return
        
        employee_id = employee_option.split(' - ')[0]
        project_id = project_option.split(' - ')[0]
        
        # 验证必填字段
        if not start_date:
            messagebox.showerror('错误', '开始日期为必填项')
            return
        
        # 日期格式验证
        try:
            if start_date:
                start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d').date()
            if end_date:
                end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d').date()
                
                # 验证结束日期是否晚于开始日期
                if end_date < start_date:
                    messagebox.showerror('错误', '结束日期不能早于开始日期')
                    return
        except ValueError:
            messagebox.showerror('错误', '日期格式无效，请使用YYYY-MM-DD格式')
            return
        
        # 添加参与记录
        if database.add_participation(employee_id, project_id, role, start_date, end_date):
            messagebox.showinfo('成功', '参与记录已添加')
            self.clear_participation_form()
            self.load_participation_data()
    
    def update_participation(self):
        # 获取表单数据
        employee_option = self.part_employee_box.get()
        project_option = self.part_project_box.get()
        role = self.part_role_box.get()
        start_date = self.part_start_entry.get() or None
        end_date = self.part_end_entry.get() or None
        
        # 获取员工ID和项目ID
        if employee_option == "无员工" or project_option == "无项目":
            messagebox.showerror('错误', '请选择有效的员工和项目')
            return
        
        employee_id = employee_option.split(' - ')[0]
        project_id = project_option.split(' - ')[0]
        
        # 验证必填字段
        if not start_date:
            messagebox.showerror('错误', '开始日期为必填项')
            return
        
        # 日期格式验证
        try:
            if start_date:
                start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d').date()
            if end_date:
                end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d').date()
                
                # 验证结束日期是否晚于开始日期
                if end_date < start_date:
                    messagebox.showerror('错误', '结束日期不能早于开始日期')
                    return
        except ValueError:
            messagebox.showerror('错误', '日期格式无效，请使用YYYY-MM-DD格式')
            return
        
        # 更新参与记录
        if database.update_participation(employee_id, project_id, role, start_date, end_date):
            messagebox.showinfo('成功', '参与记录已更新')
            self.load_participation_data()
    
    def delete_participation(self):
        # 获取表单数据
        employee_option = self.part_employee_box.get()
        project_option = self.part_project_box.get()
        
        # 获取员工ID和项目ID
        if employee_option == "无员工" or project_option == "无项目":
            messagebox.showerror('错误', '请选择有效的员工和项目')
            return
        
        employee_id = employee_option.split(' - ')[0]
        project_id = project_option.split(' - ')[0]
        
        # 确认删除
        confirm = messagebox.askyesno('确认', '确定要删除此参与记录吗？此操作不可撤销。')
        if confirm:
            if database.delete_participation(employee_id, project_id):
                messagebox.showinfo('成功', '参与记录已删除')
                self.clear_participation_form()
                self.load_participation_data()
    
    def search_participation(self):
        # 获取搜索条件
        search_by = self.part_search_box.get()
        search_value = self.part_search_entry.get()
        
        if search_by == '搜索条件' or not search_value:
            messagebox.showerror('错误', '请选择搜索条件并输入搜索值')
            return
        
        # 清空表格
        for item in self.part_tree.get_children():
            self.part_tree.delete(item)
        
        # 获取所有参与记录
        participations = database.fetch_participations()
        
        # 根据搜索条件筛选
        filtered_participations = []
        for part in participations:
            if search_by == 'Employee' and (search_value.lower() in str(part[0]).lower() or search_value.lower() in str(part[1]).lower()):
                filtered_participations.append(part)
            elif search_by == 'Project' and (search_value.lower() in str(part[2]).lower() or search_value.lower() in str(part[3]).lower()):
                filtered_participations.append(part)
            elif search_by == 'Role' and part[4] and search_value.lower() in str(part[4]).lower():
                filtered_participations.append(part)
        
        # 填充表格
        for part in filtered_participations:
            # 格式化日期
            start_date = part[5].strftime('%Y-%m-%d') if part[5] else ''
            end_date = part[6].strftime('%Y-%m-%d') if part[6] else ''
            
            self.part_tree.insert('', 'end', values=(part[0], part[1], part[2], part[3], part[4], start_date, end_date))
    
    # 统计分析选项卡
    def show_stats_tab(self):
        global current_tab
        
        # 清空内容框架
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        current_tab = "stats"
        
        # 创建统计分析界面
        stats_frame = CTkFrame(self.content_frame, fg_color='#161C30')
        stats_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 标题
        CTkLabel(stats_frame, text='统计分析', font=('arial', 24, 'bold'), text_color='white').pack()
        
        # 创建选项卡
        stats_tab_view = CTkTabview(stats_frame, width=1100, height=500)
        stats_tab_view.pack(padx=10, pady=10)
        
        # 添加选项卡
        dept_stats_tab = stats_tab_view.add("部门统计")
        project_stats_tab = stats_tab_view.add("项目统计")
        employee_stats_tab = stats_tab_view.add("员工分析")
        advanced_query_tab = stats_tab_view.add("高级查询")
        
        # 部门统计选项卡
        self.create_dept_stats_tab(dept_stats_tab)
        
        # 项目统计选项卡
        self.create_project_stats_tab(project_stats_tab)
        
        # 员工分析选项卡
        self.create_employee_stats_tab(employee_stats_tab)
        
        # 高级查询选项卡
        self.create_advanced_query_tab(advanced_query_tab)
    
    def create_dept_stats_tab(self, parent):
        # 部门统计表格
        columns = ('部门ID', '部门名称', '员工人数', '部门经理')
        
        dept_stats_tree = ttk.Treeview(parent, columns=columns, show='headings', height=15)
        for col in columns:
            dept_stats_tree.heading(col, text=col)
            if col in ['部门名称']:
                dept_stats_tree.column(col, width=200)
            else:
                dept_stats_tree.column(col, width=150)
        
        dept_stats_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=dept_stats_tree.yview)
        scrollbar.place(relx=1, rely=0.5, relheight=0.9, anchor='e')
        dept_stats_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取部门统计数据
        dept_stats = database.get_department_statistics()
        
        # 填充表格
        for dept in dept_stats:
            dept_stats_tree.insert('', 'end', values=(dept[0], dept[1], dept[2], dept[3] or "未指定"))
        
        # 添加刷新按钮
        refresh_btn = CTkButton(parent, text='刷新数据', command=lambda: self.refresh_dept_stats(dept_stats_tree))
        refresh_btn.pack(pady=10)
    
    def refresh_dept_stats(self, tree):
        # 清空表格
        for item in tree.get_children():
            tree.delete(item)
        
        # 获取部门统计数据
        dept_stats = database.get_department_statistics()
        
        # 填充表格
        for dept in dept_stats:
            tree.insert('', 'end', values=(dept[0], dept[1], dept[2], dept[3] or "未指定"))
    
    def create_project_stats_tab(self, parent):
        # 项目统计表格
        columns = ('项目ID', '项目名称', '项目状态', '成员人数', '项目负责人')
        
        proj_stats_tree = ttk.Treeview(parent, columns=columns, show='headings', height=15)
        for col in columns:
            proj_stats_tree.heading(col, text=col)
            if col in ['项目名称']:
                proj_stats_tree.column(col, width=200)
            else:
                proj_stats_tree.column(col, width=150)
        
        proj_stats_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=proj_stats_tree.yview)
        scrollbar.place(relx=1, rely=0.5, relheight=0.9, anchor='e')
        proj_stats_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取项目统计数据
        proj_stats = database.get_project_statistics()
        
        # 填充表格
        for proj in proj_stats:
            proj_stats_tree.insert('', 'end', values=(proj[0], proj[1], proj[2], proj[3], proj[4] or "未指定"))
        
        # 添加刷新按钮
        refresh_btn = CTkButton(parent, text='刷新数据', command=lambda: self.refresh_proj_stats(proj_stats_tree))
        refresh_btn.pack(pady=10)
    
    def refresh_proj_stats(self, tree):
        # 清空表格
        for item in tree.get_children():
            tree.delete(item)
        
        # 获取项目统计数据
        proj_stats = database.get_project_statistics()
        
        # 填充表格
        for proj in proj_stats:
            tree.insert('', 'end', values=(proj[0], proj[1], proj[2], proj[3], proj[4] or "未指定"))
    
    def create_employee_stats_tab(self, parent):
        # 员工分析查询框架
        query_frame = CTkFrame(parent)
        query_frame.pack(fill='x', padx=10, pady=10)
        
        # 布局使用网格结构并居中对齐
        query_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)  # 使列平均分布
        
        # 查询选项
        CTkLabel(query_frame, text='查询选项:', font=('arial', 14, 'bold')).grid(row=0, column=0, padx=10, pady=10, sticky='e')
        
        options = ['按部门查看员工分布', '按职位查看员工分布', '按项目参与情况查看']
        self.emp_query_box = CTkComboBox(query_frame, values=options, width=250, state='readonly')
        self.emp_query_box.grid(row=0, column=1, padx=10, pady=10, sticky='ew')
        self.emp_query_box.set(options[0])
        
        # 部门/职位/项目选择
        CTkLabel(query_frame, text='选择:', font=('arial', 14, 'bold')).grid(row=0, column=2, padx=10, pady=10, sticky='e')
        
        # 初始加载部门列表
        self.load_departments()
        self.emp_selection_box = CTkComboBox(query_frame, values=self.department_options, width=250, state='readonly')
        self.emp_selection_box.grid(row=0, column=3, padx=10, pady=10, sticky='ew')
        if self.department_options:
            self.emp_selection_box.set(self.department_options[0])
        
        # 查询按钮
        query_btn = CTkButton(query_frame, text='查询', width=100, command=self.execute_employee_query)
        query_btn.grid(row=0, column=4, padx=10, pady=10)
        
        # 选项变更时更新下拉列表
        self.emp_query_box.configure(command=self.update_employee_query_selection)
        
        # 结果显示框架 - 减少上边距以使结果显示更靠上
        self.emp_result_frame = CTkFrame(parent)
        self.emp_result_frame.pack(fill='both', expand=True, padx=10, pady=(5, 10))
        
        # 默认显示部门员工分布
        self.show_department_distribution()
    
    def update_employee_query_selection(self, choice):
        # 根据查询选项更新选择下拉框
        if choice == '按部门查看员工分布':
            self.load_departments()
            self.emp_selection_box.configure(values=self.department_options)
            if self.department_options:
                self.emp_selection_box.set(self.department_options[0])
        elif choice == '按职位查看员工分布':
            self.load_positions()
            self.emp_selection_box.configure(values=self.position_options)
            if self.position_options:
                self.emp_selection_box.set(self.position_options[0])
        elif choice == '按项目参与情况查看':
            self.load_projects_for_participation()
            self.emp_selection_box.configure(values=self.part_project_options)
            if self.part_project_options:
                self.emp_selection_box.set(self.part_project_options[0])
    
    def execute_employee_query(self):
        # 清空结果框架
        for widget in self.emp_result_frame.winfo_children():
            widget.destroy()
        
        query_type = self.emp_query_box.get()
        selection = self.emp_selection_box.get()
        
        if query_type == '按部门查看员工分布':
            self.show_department_employees_result(selection)
        elif query_type == '按职位查看员工分布':
            self.show_position_employees_result(selection)
        elif query_type == '按项目参与情况查看':
            self.show_project_members_result(selection)
    
    def show_department_distribution(self):
        # 清空结果框架
        for widget in self.emp_result_frame.winfo_children():
            widget.destroy()
        
        # 显示部门员工分布表格
        CTkLabel(self.emp_result_frame, text='部门员工分布', font=('arial', 18, 'bold')).pack(pady=10)
        
        # 创建表格
        columns = ('部门ID', '部门名称', '员工人数', '部门经理')
        
        dept_tree = ttk.Treeview(self.emp_result_frame, columns=columns, show='headings', height=12)
        for col in columns:
            dept_tree.heading(col, text=col)
            dept_tree.column(col, width=150)
        
        dept_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.emp_result_frame, orient='vertical', command=dept_tree.yview)
        scrollbar.place(relx=1, rely=0.5, relheight=0.8, anchor='e')
        dept_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取部门统计数据
        dept_stats = database.get_department_statistics()
        
        # 填充表格
        for dept in dept_stats:
            dept_tree.insert('', 'end', values=(dept[0], dept[1], dept[2], dept[3] or "未指定"))
    
    def show_department_employees_result(self, selection):
        if not selection or selection == "无部门":
            messagebox.showerror('错误', '请选择一个有效的部门')
            return
        
        dept_id = selection.split(' - ')[0]
        dept_name = selection.split(' - ')[1]
        
        # 显示部门员工表格
        CTkLabel(self.emp_result_frame, text=f'{dept_name} 部门员工', font=('arial', 18, 'bold')).pack(pady=10)
        
        # 创建表格
        columns = ('员工ID', '姓名', '性别', '职位')
        
        emp_tree = ttk.Treeview(self.emp_result_frame, columns=columns, show='headings', height=12)
        for col in columns:
            emp_tree.heading(col, text=col)
            emp_tree.column(col, width=150)
        
        emp_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.emp_result_frame, orient='vertical', command=emp_tree.yview)
        scrollbar.place(relx=1, rely=0.5, relheight=0.8, anchor='e')
        emp_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取部门员工数据
        employees = database.get_department_employees(dept_id)
        
        # 填充表格
        for emp in employees:
            emp_tree.insert('', 'end', values=(emp[0], emp[1], emp[2], emp[3] or "未指定"))
    
    def show_position_employees_result(self, selection):
        if not selection or selection == "无职位":
            messagebox.showerror('错误', '请选择一个有效的职位')
            return
        
        pos_id = selection.split(' - ')[0]
        pos_name = selection.split(' - ')[1]
        
        # 显示职位员工表格
        CTkLabel(self.emp_result_frame, text=f'{pos_name} 职位员工', font=('arial', 18, 'bold')).pack(pady=10)
        
        # 创建表格
        columns = ('员工ID', '姓名', '性别', '部门')
        
        emp_tree = ttk.Treeview(self.emp_result_frame, columns=columns, show='headings', height=12)
        for col in columns:
            emp_tree.heading(col, text=col)
            emp_tree.column(col, width=150)
        
        emp_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.emp_result_frame, orient='vertical', command=emp_tree.yview)
        scrollbar.place(relx=1, rely=0.5, relheight=0.8, anchor='e')
        emp_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取职位员工数据
        employees = database.get_position_employees(pos_id)
        
        # 填充表格
        for emp in employees:
            emp_tree.insert('', 'end', values=(emp[0], emp[1], emp[2], emp[3] or "未指定"))
    
    def show_project_members_result(self, selection):
        if not selection or selection == "无项目":
            messagebox.showerror('错误', '请选择一个有效的项目')
            return
        
        proj_id = selection.split(' - ')[0]
        proj_name = selection.split(' - ')[1]
        
        # 显示项目成员表格
        CTkLabel(self.emp_result_frame, text=f'{proj_name} 项目成员', font=('arial', 18, 'bold')).pack(pady=10)
        
        # 创建表格
        columns = ('员工ID', '姓名', '项目角色', '开始日期', '结束日期')
        
        members_tree = ttk.Treeview(self.emp_result_frame, columns=columns, show='headings', height=12)
        for col in columns:
            members_tree.heading(col, text=col)
            members_tree.column(col, width=150)
        
        members_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.emp_result_frame, orient='vertical', command=members_tree.yview)
        scrollbar.place(relx=1, rely=0.5, relheight=0.8, anchor='e')
        members_tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取项目成员数据
        members = database.fetch_project_members(proj_id)
        
        # 填充表格
        for member in members:
            # 格式化日期
            start_date = member[3].strftime('%Y-%m-%d') if member[3] else ''
            end_date = member[4].strftime('%Y-%m-%d') if member[4] else ''
            
            members_tree.insert('', 'end', values=(member[0], member[1], member[2], start_date, end_date))
    
    def create_advanced_query_tab(self, parent):
        # 高级查询界面
        query_frame = CTkFrame(parent)
        query_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 使用列权重使布局居中
        query_frame.grid_columnconfigure(0, weight=1)
        query_frame.grid_columnconfigure(1, weight=1)
        
        # 创建查询选择部分的容器框架
        select_frame = CTkFrame(query_frame, fg_color="transparent")
        select_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        
        # 居中布局选择框架
        select_frame.grid_columnconfigure(0, weight=1)
        select_frame.grid_columnconfigure(1, weight=2)
        select_frame.grid_columnconfigure(2, weight=1)
        
        # 查询选项标签和下拉框
        CTkLabel(select_frame, text='选择查询:', font=('arial', 16, 'bold')).grid(row=0, column=0,padx=(80, 0), sticky='e')
        
        queries = [
            '查询同时参与多个项目的员工',
            '查询未分配部门的员工',
            '查询各部门平均工资',
            '查询项目人数最多的前三个项目',
            '按关键字搜索员工（姓名、部门、职位）'
        ]
        
        self.advanced_query_box = CTkComboBox(select_frame, values=queries, width=350, state='readonly')
        self.advanced_query_box.grid(row=0, column=1, padx=(0, 110))
        self.advanced_query_box.set(queries[0])
        
        # 搜索关键字输入框和标签（居中布局）
        keyword_frame = CTkFrame(query_frame, fg_color="transparent")
        keyword_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky='ew')
        
        keyword_frame.grid_columnconfigure(0, weight=1)
        keyword_frame.grid_columnconfigure(1, weight=1)
        keyword_frame.grid_columnconfigure(2, weight=1)
        
        self.keyword_label = CTkLabel(keyword_frame, text='关键字:', font=('arial', 14))
        self.keyword_label.grid(row=0, column=0, padx=10, sticky='e')
        
        self.keyword_entry = CTkEntry(keyword_frame, width=250)
        self.keyword_entry.grid(row=0, column=1, padx=10, sticky='ew')
        
        # 执行查询按钮
        execute_btn = CTkButton(keyword_frame, text='执行查询', width=120, command=self.execute_advanced_query)
        execute_btn.grid(row=0, column=2, padx=10, pady=5, sticky='w')
        
        # 设置命令以在选择更改时更新界面
        self.advanced_query_box.configure(command=self.update_advanced_query_ui)
        
        # 结果显示框架
        result_frame = CTkFrame(query_frame)
        result_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=(5, 10), sticky='nsew')
        
        # 配置行和列权重，使结果框架可以扩展
        query_frame.grid_rowconfigure(2, weight=1)
        
        # 创建结果显示区
        self.result_label = CTkLabel(result_frame, text='查询结果', font=('arial', 16, 'bold'))
        self.result_label.pack(pady=(5, 5))  # 减少上下边距
        
        # 创建滚动框架
        self.result_scroll = CTkScrollableFrame(result_frame, width=700, height=350)
        self.result_scroll.pack(fill='both', expand=True, padx=10)
    
    def update_advanced_query_ui(self, choice):
        # 根据选择的查询类型更新界面
        if choice == '按关键字搜索员工（姓名、部门、职位）':
            self.keyword_label.grid(row=0, column=0, padx=10, pady=5, sticky='e')
            self.keyword_entry.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        else:
            self.keyword_label.grid_forget()
            self.keyword_entry.grid_forget()
    
    def execute_advanced_query(self):
        # 清空结果区
        for widget in self.result_scroll.winfo_children():
            widget.destroy()
        
        query_type = self.advanced_query_box.get()
        
        if query_type == '查询同时参与多个项目的员工':
            self.query_multi_project_employees()
        elif query_type == '查询未分配部门的员工':
            self.query_employees_without_dept()
        elif query_type == '查询各部门平均工资':
            self.query_dept_avg_salary()
        elif query_type == '查询项目人数最多的前三个项目':
            self.query_top_projects_by_members()
        elif query_type == '按关键字搜索员工（姓名、部门、职位）':
            keyword = self.keyword_entry.get()
            if not keyword:
                messagebox.showerror('错误', '请输入搜索关键字')
                return
            self.query_employees_by_keyword(keyword)
    
    def query_multi_project_employees(self):
        # 查询同时参与多个项目的员工
        # 这需要在数据库模块中添加新的查询函数
        
        # 创建表格来显示结果
        columns = ('员工ID', '姓名', '参与项目数量')
        
        result_tree = ttk.Treeview(self.result_scroll, columns=columns, show='headings', height=10)
        for col in columns:
            result_tree.heading(col, text=col)
            result_tree.column(col, width=200)
        
        result_tree.pack(fill='both', expand=True)
        
        # 模拟数据（实际应从数据库获取）
        # 在真实实现中，应该添加一个数据库函数来获取这些数据
        results = [
            ('E001', '张三', 2),
            ('E002', '李四', 2),
            ('E003', '王五', 2)
        ]
        
        # 填充表格
        for result in results:
            result_tree.insert('', 'end', values=result)
    
    def query_employees_without_dept(self):
        # 查询未分配部门的员工
        # 这需要在数据库模块中添加新的查询函数
        
        # 创建表格来显示结果
        columns = ('员工ID', '姓名', '性别', '联系方式')
        
        result_tree = ttk.Treeview(self.result_scroll, columns=columns, show='headings', height=10)
        for col in columns:
            result_tree.heading(col, text=col)
            result_tree.column(col, width=150)
        
        result_tree.pack(fill='both', expand=True)
        
        # 模拟数据（实际应从数据库获取）
        # 在真实实现中，应该添加一个数据库函数来获取这些数据
        # 因为我们的示例数据中没有未分配部门的员工，所以这里显示为空
        CTkLabel(self.result_scroll, text='没有找到未分配部门的员工').pack(pady=20)
    
    def query_dept_avg_salary(self):
        # 查询各部门平均工资
        # 这需要在数据库模块中添加新的查询函数
        
        # 创建表格来显示结果
        columns = ('部门ID', '部门名称', '平均工资')
        
        result_tree = ttk.Treeview(self.result_scroll, columns=columns, show='headings', height=10)
        for col in columns:
            result_tree.heading(col, text=col)
            result_tree.column(col, width=200)
        
        result_tree.pack(fill='both', expand=True)
        
        # 模拟数据（实际应从数据库获取）
        # 在真实实现中，应该添加一个数据库函数来获取这些数据
        results = [
            ('D001', '研发部', 14333.33),
            ('D002', '市场部', 20000.00),
            ('D003', '人事部', 0.00)  # 没有员工
        ]
        
        # 填充表格
        for result in results:
            result_tree.insert('', 'end', values=result)
    
    def query_top_projects_by_members(self):
        # 查询项目人数最多的前三个项目
        # 这需要在数据库模块中添加新的查询函数
        
        # 创建表格来显示结果
        columns = ('项目ID', '项目名称', '成员数量', '项目状态')
        
        result_tree = ttk.Treeview(self.result_scroll, columns=columns, show='headings', height=10)
        for col in columns:
            result_tree.heading(col, text=col)
            result_tree.column(col, width=150)
        
        result_tree.pack(fill='both', expand=True)
        
        # 获取项目统计数据
        proj_stats = database.get_project_statistics()
        
        # 按成员数量排序并取前三名
        sorted_projects = sorted(proj_stats, key=lambda x: x[3], reverse=True)[:3]
        
        # 填充表格
        for proj in sorted_projects:
            result_tree.insert('', 'end', values=(proj[0], proj[1], proj[3], proj[2]))
    
    def query_employees_by_keyword(self, keyword):
        # 按关键字搜索员工
        
        # 创建表格来显示结果
        columns = ('员工ID', '姓名', '性别', '联系方式', '部门', '职位')
        
        result_tree = ttk.Treeview(self.result_scroll, columns=columns, show='headings', height=10)
        for col in columns:
            result_tree.heading(col, text=col)
            if col in ['联系方式']:
                result_tree.column(col, width=150)
            else:
                result_tree.column(col, width=100)
        
        result_tree.pack(fill='both', expand=True)
        #result_tree.grid(row=1)
        # 获取搜索结果
        employees = database.search_employees_by_keyword(keyword)
        
        # 填充表格
        if employees:
            for emp in employees:
                result_tree.insert('', 'end', values=(emp[0], emp[1], emp[2], emp[3], emp[4], emp[5]))
        else:
            CTkLabel(self.result_scroll, text=f'没有找到匹配关键字"{keyword}"的员工').pack(pady=20)

# 主程序
if __name__ == "__main__":
    # 在直接运行此文件时执行的代码
    # 初始化数据库连接
    import database
    
    try:
        # 连接数据库
        database.connect_database()
        
        # 检查是否已存在数据，如果不存在则添加样例数据
        mycursor = database.mycursor
        mycursor.execute("SELECT COUNT(*) FROM Department")
        dept_count = mycursor.fetchone()[0]
        
        if dept_count == 0:
            database.insert_sample_data()
    except Exception as e:
        from tkinter import messagebox
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror('错误', f'初始化数据库时出错: {str(e)}')
        root.destroy()
        import sys
        sys.exit(1)
    
    # 启动主界面
    root = CTk()
    app = EmployeeManagementSystem(root)
    root.mainloop() 