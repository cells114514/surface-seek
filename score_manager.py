import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox, ttk



# 设置文件路径为当前脚本所在目录
#script_dir = os.path.dirname(os.path.abspath(__file__))
#excel_file = os.path.join(script_dir, 'scores.xlsx')
#print(f"Excel文件路径: {excel_file}")



class excel_methods():
    def __init__(self, file_name='default.xlsx', sheet_name='default'):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_file = os.path.join(self.script_dir, file_name)
        self.sheet_name = sheet_name
        self.ensure_exists(self.sheet_name)

    def ensure_exists(self, sheet_name):          #确保文件和sheet存在
        """
        确保指定的Excel文件和sheet存在，如果不存在则创建。
        """

        if not os.path.exists(self.excel_file):
            # 文件不存在，创建新的Excel文件
            with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
                pd.DataFrame(columns=['']).to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"文件不存在：已创建了新文件: {self.excel_file}")
        else:
            # 文件存在，检查sheet是否存在
            try:
                pd.read_excel(self.excel_file, sheet_name=sheet_name)
            except ValueError:
                # sheet不存在，创建新的sheet
                with pd.ExcelWriter(self.excel_file, engine='openpyxl', mode='a') as writer:
                    pd.DataFrame(columns=['']).to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"工作表不存在: 已在现有文件中创建了新工作表: {sheet_name}")


    def rewrite(self, temp_dataframe, sheet):
        with pd.ExcelWriter(self.excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            temp_dataframe.to_excel(writer, sheet_name=sheet, index=False)
            # 将更新后的DataFrame写回到Excel文件


    def show_sheet(self, sheet):                  #读取指定sheet
        self.ensure_exists(sheet)
        df = pd.read_excel(self.excel_file, sheet_name=sheet)
        print(df)



    def append_row(self, sheet, **kwargs):           #表尾加一行
        # ensure_exists(sheet)
        df = pd.read_excel(self.excel_file, sheet_name=sheet)
        new_row = pd.DataFrame([kwargs])
        df = pd.concat([df, new_row], ignore_index=True)

        self.rewrite(df, sheet)
        print(f"append: 已添加数据到工作表 {sheet}: {kwargs}")


    def insert_a_row(self, sheet, index, **kwargs):  #在指定行插入一行
        self.ensure_exists(sheet)
        df = pd.read_excel(self.excel_file, sheet_name=sheet)
        new_row = pd.DataFrame([kwargs])
        if index < 0 or index > len(df):
            print(f"insert: 索引 {index} 超出范围，无法插入行。")
            return None
        else:
            df = pd.concat([df.iloc[:index], new_row, df.iloc[index:]], ignore_index=True)
            self.rewrite(df, sheet)
            print(f"insert: 已在工作表 {sheet} 的索引 {index} 处插入新行: {kwargs}")
            return df
        



    def find_row(self, sheet, col, name):         #用任意一列的值查找
        self.ensure_exists(sheet)
        df = pd.read_excel(self.excel_file, sheet_name=sheet)
        if col not in df.columns:
            print(f"find: 列 '{col}' 在工作表 '{sheet}' 中不存在。")
            return None

        result = df[df[col] == name]
        if result.empty:
            print(f"find: 未找到 '{name}' 在工作表 '{sheet}' 中。")
            return None
        else:
            
            idp = df.index[df[col] == name].tolist()
            
            row_list = []
            for idx in idp:
                row_list.append(df.iloc[idx].to_dict())
                print(f"第{idx}行数据: {df.iloc[idx].to_dict()}")
            
            return pd.DataFrame(row_list, index=idp)


    def delete_one_row(self, sheet, col, name):       #删除指定的一行
        self.ensure_exists(sheet)
        df = pd.read_excel(self.excel_file, sheet_name=sheet)
        target_row = self.find_row(sheet, col, name)  # 提取符合条件的行
        if target_row is None:
            print(f"delete: 未找到 '{name}' 在工作表 '{sheet}' 中。")
            return None
        while True:
            try:
                indices_to_drop = list(map(int, input("delete: 输入希望删除的行号索引（多个用空格分隔）: ").split()))
                if not all(idx in target_row.index for idx in indices_to_drop):
                    print("delete: 输入的行号索引无效，请重新输入。")
                    continue
                break
            except ValueError:
                print("delete: 输入的行号索引无效，请输入整数。")
                return None
        
        for idx in sorted(indices_to_drop, reverse=True):
                df.drop(index=idx, inplace=True)


        self.rewrite(df, sheet)
        if len(indices_to_drop) == 0:
            print(f"delete: 未删除任何行，因为输入的索引为空。")
        return indices_to_drop



    def delete_range(self, sheet, col, name):       #删除同名的全部行
        self.ensure_exists(sheet)
        df = pd.read_excel(self.excel_file, sheet_name=sheet)

        # 检查列是否存在
        if col not in df.columns:
            print(f"列 '{col}' 在工作表 '{sheet}' 中不存在。")
            return None
        
        # 获取删除前的行数
        original_count = len(df)
        
        # 保留不匹配的行（相当于删除匹配的行）
        df = df[df[col] != name].reset_index(drop=True)
        delete_length = original_count - len(df)
        # 检查是否有行被删除
        if len(df) == original_count: 
            print(f"未找到 '{name}' 在工作表 '{sheet}' 中。")
            return 0

        self.rewrite(df, sheet)
        print(f"已删除 '{name}' 从工作表 '{sheet}'")
        return delete_length


    def edit_row(self, sheet, col, name):     #unfinished(input)
        edit_index = self.delete_one_row(sheet, col, name)
        for idx in edit_index:
            input_data = input(f"edit: 输入第{idx}行新的数据（格式为表头=数据，不同列之间以逗号分隔）: ")
            # 这里应该解析输入的键值对，而不是使用 dir 函数
            pairs = input_data.split(',')
            new_data = {}
            for pair in pairs:
                key, value = pair.split('=')
                new_data[key.strip()] = value.strip()
            self.insert_a_row(sheet, idx, **new_data)

     


# score = excel_methods(file_name='scores.xlsx', sheet_name='sheet1')  # 创建实例，指定文件名和工作表名
# score.append_row('sheet1', Name='Alice', sex = 'Female', Score=90)  # 添加一行数据




class ScoreManagerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("成绩管理系统")
        self.root.geometry("800x600")
        
        # 创建 excel_methods 实例
        self.score_manager = excel_methods(file_name='scores.xlsx', sheet_name='sheet1')
        
        self.create_widgets()
    
    def create_widgets(self):
        # 创建主界面布局
        self.create_menu()
        self.create_main_frame()
        
    def create_menu(self):
        # 创建菜单栏（窗口上方的横条）
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="新建", command=self.new_file)
        file_menu.add_command(label="打开", command=self.open_file)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
    def create_main_frame(self):
        # 主操作按钮
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="添加成绩", command=self.add_score, 
                 width=15, height=2).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="查看成绩", command=self.view_scores, 
                 width=15, height=2).grid(row=0, column=1, padx=10)
        tk.Button(btn_frame, text="编辑成绩", command=self.edit_score, 
                 width=15, height=2).grid(row=0, column=2, padx=10)
        tk.Button(btn_frame, text="删除成绩", command=self.delete_score, 
                 width=15, height=2).grid(row=1, column=0, padx=10, pady=10)
        tk.Button(btn_frame, text="查找学生", command=self.find_student, 
                 width=15, height=2).grid(row=1, column=1, padx=10, pady=10)
        
        # 创建数据显示区域
        self.create_data_display()
    
    def create_data_display(self):
        # 创建表格显示区域
        display_frame = tk.Frame(self.root)
        display_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 使用 Treeview 创建表格
        columns = ('索引', '姓名', '性别', '成绩')
        self.tree = ttk.Treeview(display_frame, columns=columns, show='headings')
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(display_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def add_score(self):
        # 创建添加成绩的对话框
        self.create_input_dialog("添加成绩", self.handle_add_score)
    
    def create_input_dialog(self, title, callback):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 输入字段
        tk.Label(dialog, text="姓名:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        name_entry = tk.Entry(dialog)
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        
        tk.Label(dialog, text="性别:").grid(row=1, column=0, padx=10, pady=5, sticky='w')
        sex_entry = tk.Entry(dialog)
        sex_entry.grid(row=1, column=1, padx=10, pady=5)
        
        tk.Label(dialog, text="成绩:").grid(row=2, column=0, padx=10, pady=5, sticky='w')
        score_entry = tk.Entry(dialog)
        score_entry.grid(row=2, column=1, padx=10, pady=5)
        
        # 按钮
        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        tk.Button(btn_frame, text="确定", 
                 command=lambda: callback(name_entry.get(), sex_entry.get(), score_entry.get(), dialog)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def handle_add_score(self, name, sex, score, dialog):
        try:
            if name and sex and score:
                self.score_manager.append_row('sheet1', Name=name, sex=sex, Score=int(score))
                messagebox.showinfo("成功", "成绩添加成功！")
                self.refresh_display()
                dialog.destroy()
            else:
                messagebox.showerror("错误", "请填写所有字段！")
        except ValueError:
            messagebox.showerror("错误", "成绩必须是数字！")
    
    def view_scores(self):
        self.refresh_display()
    
    def refresh_display(self):
        # 清空当前显示
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 读取数据并显示
        try:
            df = pd.read_excel(self.score_manager.excel_file, sheet_name='sheet1')
            for idx, row in df.iterrows():
                self.tree.insert('', 'end', values=(idx, row.get('Name', ''), row.get('sex', ''), row.get('Score', '')))
        except Exception as e:
            messagebox.showerror("错误", f"读取数据失败: {str(e)}")
    
    def edit_score(self):
        messagebox.showinfo("提示", "编辑功能待实现")
    
    def delete_score(self):
        messagebox.showinfo("提示", "删除功能待实现")
    
    def find_student(self):
        messagebox.showinfo("提示", "查找功能待实现")
    
    def new_file(self):
        # 创建新文件
        input_box = tk.Toplevel(self.root)
        input_box.title("新建文件")
        input_box.geometry("300x150")
        input_box.transient(self.root)      # 设置为顶级窗口
        input_box.grab_set()                # 强制优先窗口
        # self.root.wait_window(input_box)    # 等待窗口关闭

        tk.Label(input_box, text="新文件名:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        file_entry = tk.Entry(input_box)
        file_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(input_box, text="工作表名:").grid(row=1, column=0, padx=10, pady=5, sticky='w')
        sheet_entry = tk.Entry(input_box)
        sheet_entry.grid(row=1, column=1, padx=10, pady=5)


        def create():
            if not file_entry.get() or not sheet_entry.get():
                messagebox.showerror("错误", "文件名和工作表名不能为空！")
                return
   
            # 创建新的 Excel 文件和工作表
            excel_methods(file_name=file_entry.get() + '.xlsx', sheet_name=sheet_entry.get()).ensure_exists(sheet_entry.get())
            messagebox.showinfo("成功", f"已创建新文件: {file_entry.get()}.xlsx")
            input_box.destroy()
            self.refresh_display()

        btn_btm = tk.Frame(input_box)
        btn_btm.grid(row=3, column=0, columnspan=2, pady=10)
        
        tk.Button(btn_btm, text="取消", command=input_box.destroy).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_btm, text="创建", command=create).pack(side=tk.LEFT, padx=5)



    
    def open_file(self):
        os.startfile(self.score_manager.excel_file)
        
    def run(self):
        self.root.mainloop()

# 使用示例
if __name__ == "__main__":
    app = ScoreManagerGUI()
    app.run()
