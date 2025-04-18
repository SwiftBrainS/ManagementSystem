import pymysql
from tkinter import messagebox
import datetime

def connect_database():
    global mycursor, conn
    try:
        conn = pymysql.connect(host='localhost', user='root', password='143321')
        mycursor = conn.cursor()
    except:
        messagebox.showerror('Error', '数据库连接失败')
        return

    # 创建数据库
    mycursor.execute('CREATE DATABASE IF NOT EXISTS employee_database')
    mycursor.execute('USE employee_database')
    
    # 表结构定义
    try:
        # 创建部门表 (无外键)
        mycursor.execute('''CREATE TABLE IF NOT EXISTS Department (
                        DeptID VARCHAR(20) PRIMARY KEY,
                        DeptName VARCHAR(50) NOT NULL,
                        ManagerID VARCHAR(20),
                        EstablishDate DATE,
                        Description TEXT
                        )''')
        
        # 创建职位表
        mycursor.execute('''CREATE TABLE IF NOT EXISTS `Position` (
                        PositionID VARCHAR(20) PRIMARY KEY,
                        PositionName VARCHAR(50) NOT NULL,
                        PositionCategory VARCHAR(50),
                        PositionLevel INT,
                        BaseSalaryGrade DECIMAL(10, 2)
                        )''')
        
        # 创建员工表
        mycursor.execute('''CREATE TABLE IF NOT EXISTS Employee (
                        EmployeeID VARCHAR(20) PRIMARY KEY,
                        Name VARCHAR(50) NOT NULL,
                        Gender VARCHAR(20),
                        BirthDate DATE,
                        ContactInfo VARCHAR(100),
                        HireDate DATE,
                        DeptID VARCHAR(20),
                        PositionID VARCHAR(20),
                        FOREIGN KEY (DeptID) REFERENCES Department(DeptID) ON DELETE SET NULL,
                        FOREIGN KEY (PositionID) REFERENCES `Position`(PositionID) ON DELETE SET NULL
                        )''')
        
        # 创建项目表
        mycursor.execute('''CREATE TABLE IF NOT EXISTS Project (
                        ProjectID VARCHAR(20) PRIMARY KEY,
                        ProjectName VARCHAR(100) NOT NULL,
                        StartDate DATE,
                        EndDate DATE,
                        Status VARCHAR(20),
                        LeaderID VARCHAR(20),
                        FOREIGN KEY (LeaderID) REFERENCES Employee(EmployeeID) ON DELETE SET NULL
                        )''')
        
        # 创建参与表
        mycursor.execute('''CREATE TABLE IF NOT EXISTS Participation (
                        EmployeeID VARCHAR(20),
                        ProjectID VARCHAR(20),
                        Role VARCHAR(50),
                        StartDate DATE,
                        EndDate DATE,
                        PRIMARY KEY (EmployeeID, ProjectID),
                        FOREIGN KEY (EmployeeID) REFERENCES Employee(EmployeeID) ON DELETE CASCADE,
                        FOREIGN KEY (ProjectID) REFERENCES Project(ProjectID) ON DELETE CASCADE
                        )''')
        
        # 尝试为Department表添加外键约束
        # 先检查是否已存在约束并尝试删除
        try:
            mycursor.execute('''ALTER TABLE Department DROP FOREIGN KEY fk_manager''')
        except:
            # 如果不存在约束，忽略错误
            pass
        
        # 添加外键约束
        mycursor.execute('''ALTER TABLE Department 
                        ADD CONSTRAINT fk_manager 
                        FOREIGN KEY (ManagerID) REFERENCES Employee(EmployeeID) ON DELETE SET NULL''')
    
        conn.commit()
    except Exception as e:
        conn.rollback()
        messagebox.showerror('错误', f'创建数据库表时出错: {str(e)}')
        raise e

# 部门相关函数
def insert_department(dept_id, dept_name, manager_id=None, establish_date=None, description=None):
    try:
        mycursor.execute('''INSERT INTO Department (DeptID, DeptName, ManagerID, EstablishDate, Description) 
                         VALUES (%s, %s, %s, %s, %s)''', 
                         (dept_id, dept_name, manager_id, establish_date, description))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'添加部门失败: {str(e)}')
        return False

def update_department(dept_id, dept_name, manager_id=None, establish_date=None, description=None):
    try:
        mycursor.execute('''UPDATE Department 
                         SET DeptName = %s, ManagerID = %s, EstablishDate = %s, Description = %s 
                         WHERE DeptID = %s''', 
                         (dept_name, manager_id, establish_date, description, dept_id))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'更新部门失败: {str(e)}')
        return False

def delete_department(dept_id):
    try:
        mycursor.execute('DELETE FROM Department WHERE DeptID = %s', (dept_id,))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'删除部门失败: {str(e)}')
        return False

def fetch_departments():
    mycursor.execute('SELECT * FROM Department')
    return mycursor.fetchall()

def dept_exists(dept_id):
    mycursor.execute('SELECT COUNT(*) FROM Department WHERE DeptID = %s', (dept_id,))
    return mycursor.fetchone()[0] > 0

# 职位相关函数
def insert_position(position_id, position_name, category=None, level=None, base_salary=None):
    try:
        mycursor.execute('''INSERT INTO Position (PositionID, PositionName, PositionCategory, PositionLevel, BaseSalaryGrade) 
                         VALUES (%s, %s, %s, %s, %s)''', 
                         (position_id, position_name, category, level, base_salary))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'添加职位失败: {str(e)}')
        return False

def update_position(position_id, position_name, category=None, level=None, base_salary=None):
    try:
        mycursor.execute('''UPDATE Position 
                         SET PositionName = %s, PositionCategory = %s, PositionLevel = %s, BaseSalaryGrade = %s 
                         WHERE PositionID = %s''', 
                         (position_name, category, level, base_salary, position_id))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'更新职位失败: {str(e)}')
        return False

def delete_position(position_id):
    try:
        mycursor.execute('DELETE FROM Position WHERE PositionID = %s', (position_id,))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'删除职位失败: {str(e)}')
        return False

def fetch_positions():
    mycursor.execute('SELECT * FROM Position')
    return mycursor.fetchall()

def position_exists(position_id):
    mycursor.execute('SELECT COUNT(*) FROM Position WHERE PositionID = %s', (position_id,))
    return mycursor.fetchone()[0] > 0

# 员工相关函数
def insert_employee(employee_id, name, gender=None, birth_date=None, contact_info=None, 
                   hire_date=None, dept_id=None, position_id=None):
    try:
        mycursor.execute('''INSERT INTO Employee 
                         (EmployeeID, Name, Gender, BirthDate, ContactInfo, HireDate, DeptID, PositionID) 
                         VALUES (%s, %s, %s, %s, %s, %s, %s, %s)''', 
                         (employee_id, name, gender, birth_date, contact_info, hire_date, dept_id, position_id))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'添加员工失败: {str(e)}')
        return False

def update_employee(employee_id, name, gender=None, birth_date=None, contact_info=None, 
                    hire_date=None, dept_id=None, position_id=None):
    try:
        mycursor.execute('''UPDATE Employee 
                         SET Name = %s, Gender = %s, BirthDate = %s, ContactInfo = %s, HireDate = %s, 
                         DeptID = %s, PositionID = %s 
                         WHERE EmployeeID = %s''', 
                         (name, gender, birth_date, contact_info, hire_date, dept_id, position_id, employee_id))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'更新员工失败: {str(e)}')
        return False

def delete_employee(employee_id):
    try:
        mycursor.execute('DELETE FROM Employee WHERE EmployeeID = %s', (employee_id,))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'删除员工失败: {str(e)}')
        return False

def fetch_employees():
    mycursor.execute('''SELECT e.EmployeeID, e.Name, e.Gender, e.BirthDate, e.ContactInfo, e.HireDate, 
                      d.DeptName, p.PositionName 
                      FROM Employee e 
                      LEFT JOIN Department d ON e.DeptID = d.DeptID 
                      LEFT JOIN Position p ON e.PositionID = p.PositionID''')
    return mycursor.fetchall()

def employee_exists(employee_id):
    mycursor.execute('SELECT COUNT(*) FROM Employee WHERE EmployeeID = %s', (employee_id,))
    return mycursor.fetchone()[0] > 0

# 项目相关函数
def insert_project(project_id, project_name, start_date=None, end_date=None, status=None, leader_id=None):
    try:
        mycursor.execute('''INSERT INTO Project 
                         (ProjectID, ProjectName, StartDate, EndDate, Status, LeaderID) 
                         VALUES (%s, %s, %s, %s, %s, %s)''', 
                         (project_id, project_name, start_date, end_date, status, leader_id))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'添加项目失败: {str(e)}')
        return False

def update_project(project_id, project_name, start_date=None, end_date=None, status=None, leader_id=None):
    try:
        mycursor.execute('''UPDATE Project 
                         SET ProjectName = %s, StartDate = %s, EndDate = %s, Status = %s, LeaderID = %s 
                         WHERE ProjectID = %s''', 
                         (project_name, start_date, end_date, status, leader_id, project_id))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'更新项目失败: {str(e)}')
        return False

def delete_project(project_id):
    try:
        mycursor.execute('DELETE FROM Project WHERE ProjectID = %s', (project_id,))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'删除项目失败: {str(e)}')
        return False

def fetch_projects():
    mycursor.execute('''SELECT p.ProjectID, p.ProjectName, p.StartDate, p.EndDate, p.Status, 
                      e.Name as LeaderName 
                      FROM Project p 
                      LEFT JOIN Employee e ON p.LeaderID = e.EmployeeID''')
    return mycursor.fetchall()

def project_exists(project_id):
    mycursor.execute('SELECT COUNT(*) FROM Project WHERE ProjectID = %s', (project_id,))
    return mycursor.fetchone()[0] > 0

# 参与相关函数
def add_participation(employee_id, project_id, role=None, start_date=None, end_date=None):
    try:
        mycursor.execute('''INSERT INTO Participation 
                         (EmployeeID, ProjectID, Role, StartDate, EndDate) 
                         VALUES (%s, %s, %s, %s, %s)''', 
                         (employee_id, project_id, role, start_date, end_date))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'添加参与记录失败: {str(e)}')
        return False

def update_participation(employee_id, project_id, role=None, start_date=None, end_date=None):
    try:
        mycursor.execute('''UPDATE Participation 
                         SET Role = %s, StartDate = %s, EndDate = %s 
                         WHERE EmployeeID = %s AND ProjectID = %s''', 
                         (role, start_date, end_date, employee_id, project_id))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'更新参与记录失败: {str(e)}')
        return False

def delete_participation(employee_id, project_id):
    try:
        mycursor.execute('DELETE FROM Participation WHERE EmployeeID = %s AND ProjectID = %s', 
                        (employee_id, project_id))
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'删除参与记录失败: {str(e)}')
        return False

def fetch_participations():
    mycursor.execute('''SELECT p.EmployeeID, e.Name, p.ProjectID, pr.ProjectName, 
                      p.Role, p.StartDate, p.EndDate 
                      FROM Participation p 
                      JOIN Employee e ON p.EmployeeID = e.EmployeeID 
                      JOIN Project pr ON p.ProjectID = pr.ProjectID''')
    return mycursor.fetchall()

def fetch_project_members(project_id):
    mycursor.execute('''SELECT e.EmployeeID, e.Name, p.Role, p.StartDate, p.EndDate 
                      FROM Participation p 
                      JOIN Employee e ON p.EmployeeID = e.EmployeeID 
                      WHERE p.ProjectID = %s''', (project_id,))
    return mycursor.fetchall()

def fetch_employee_projects(employee_id):
    mycursor.execute('''SELECT pr.ProjectID, pr.ProjectName, p.Role, p.StartDate, p.EndDate 
                      FROM Participation p 
                      JOIN Project pr ON p.ProjectID = pr.ProjectID 
                      WHERE p.EmployeeID = %s''', (employee_id,))
    return mycursor.fetchall()

# 多表查询功能
def get_department_employees(dept_id):
    """获取指定部门的所有员工"""
    mycursor.execute('''SELECT e.EmployeeID, e.Name, e.Gender, p.PositionName 
                      FROM Employee e 
                      LEFT JOIN Position p ON e.PositionID = p.PositionID 
                      WHERE e.DeptID = %s''', (dept_id,))
    return mycursor.fetchall()

def get_position_employees(position_id):
    """获取指定职位的所有员工"""
    mycursor.execute('''SELECT e.EmployeeID, e.Name, e.Gender, d.DeptName 
                      FROM Employee e 
                      LEFT JOIN Department d ON e.DeptID = d.DeptID 
                      WHERE e.PositionID = %s''', (position_id,))
    return mycursor.fetchall()

def get_employee_detail(employee_id):
    """获取员工详细信息（包括部门和职位）"""
    mycursor.execute('''SELECT e.*, d.DeptName, p.PositionName, p.BaseSalaryGrade 
                      FROM Employee e 
                      LEFT JOIN Department d ON e.DeptID = d.DeptID 
                      LEFT JOIN Position p ON e.PositionID = p.PositionID 
                      WHERE e.EmployeeID = %s''', (employee_id,))
    return mycursor.fetchone()

def search_employees_by_keyword(keyword):
    """按关键字搜索员工"""
    search_param = f'%{keyword}%'
    mycursor.execute('''SELECT e.EmployeeID, e.Name, e.Gender, e.ContactInfo, 
                      d.DeptName, p.PositionName 
                      FROM Employee e 
                      LEFT JOIN Department d ON e.DeptID = d.DeptID 
                      LEFT JOIN Position p ON e.PositionID = p.PositionID 
                      WHERE e.Name LIKE %s 
                      OR e.EmployeeID LIKE %s 
                      OR d.DeptName LIKE %s 
                      OR p.PositionName LIKE %s''', 
                      (search_param, search_param, search_param, search_param))
    return mycursor.fetchall()

def get_department_statistics():
    """获取部门统计信息"""
    mycursor.execute('''SELECT d.DeptID, d.DeptName, COUNT(e.EmployeeID) as EmployeeCount, 
                      e2.Name as ManagerName 
                      FROM Department d 
                      LEFT JOIN Employee e ON d.DeptID = e.DeptID 
                      LEFT JOIN Employee e2 ON d.ManagerID = e2.EmployeeID 
                      GROUP BY d.DeptID''')
    return mycursor.fetchall()

def get_project_statistics():
    """获取项目统计信息"""
    mycursor.execute('''SELECT p.ProjectID, p.ProjectName, p.Status, 
                      COUNT(pa.EmployeeID) as MemberCount, 
                      e.Name as LeaderName 
                      FROM Project p 
                      LEFT JOIN Participation pa ON p.ProjectID = pa.ProjectID 
                      LEFT JOIN Employee e ON p.LeaderID = e.EmployeeID 
                      GROUP BY p.ProjectID''')
    return mycursor.fetchall()

# 清空表格
def truncate_table(table_name):
    try:
        mycursor.execute(f'TRUNCATE TABLE {table_name}')
        conn.commit()
        return True
    except pymysql.Error as e:
        conn.rollback()
        messagebox.showerror('错误', f'清空表格失败: {str(e)}')
        return False

# 向现有数据库中添加示例数据
def insert_sample_data():
    # 添加部门
    insert_department('D001', '研发部', None, '2020-01-01', '负责产品研发')
    insert_department('D002', '市场部', None, '2020-01-01', '负责市场营销')
    insert_department('D003', '人事部', None, '2020-01-01', '负责人力资源管理')
    
    # 添加职位
    insert_position('P001', '高级工程师', '技术', 3, 15000)
    insert_position('P002', '初级工程师', '技术', 1, 8000)
    insert_position('P003', '部门经理', '管理', 4, 20000)
    insert_position('P004', '市场专员', '市场', 2, 10000)
    
    # 添加员工
    insert_employee('E001', '张三', '男', '1990-05-15', '13800138001', '2020-02-01', 'D001', 'P003')
    insert_employee('E002', '李四', '男', '1992-08-20', '13800138002', '2020-03-15', 'D001', 'P001')
    insert_employee('E003', '王五', '女', '1995-11-10', '13800138003', '2021-01-10', 'D001', 'P002')
    insert_employee('E004', '赵六', '女', '1991-07-22', '13800138004', '2020-05-05', 'D002', 'P003')
    
    # 更新部门管理者
    update_department('D001', '研发部', 'E001', '2020-01-01', '负责产品研发')
    update_department('D002', '市场部', 'E004', '2020-01-01', '负责市场营销')
    
    # 添加项目
    insert_project('PRJ001', '客户管理系统开发', '2022-01-01', '2022-06-30', '已完成', 'E001')
    insert_project('PRJ002', '移动应用开发', '2022-03-15', '2022-12-31', '进行中', 'E002')
    
    # 添加参与记录
    add_participation('E001', 'PRJ001', '项目经理', '2022-01-01', '2022-06-30')
    add_participation('E002', 'PRJ001', '技术负责人', '2022-01-01', '2022-06-30')
    add_participation('E003', 'PRJ001', '开发人员', '2022-01-15', '2022-06-30')
    add_participation('E002', 'PRJ002', '项目经理', '2022-03-15', None)
    add_participation('E003', 'PRJ002', '开发人员', '2022-03-20', None)

# 初始化连接
connect_database()