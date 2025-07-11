import pandas as pd
import os

# 设置文件路径为当前脚本所在目录
script_dir = os.path.dirname(os.path.abspath(__file__))
excel_file = os.path.join(script_dir, 'scores.xlsx')
#print(f"Excel文件路径: {excel_file}")

def ensure_exists(sheet_name):          #确保文件和sheet存在
    """
    确保指定的Excel文件和sheet存在，如果不存在则创建。
    """
    
    if not os.path.exists(excel_file):
        # 文件不存在，创建新的Excel文件
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            pd.DataFrame(columns=['']).to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"文件不存在：已创建了新文件: {excel_file}")
    else:
        # 文件存在，检查sheet是否存在
        try:
            pd.read_excel(excel_file, sheet_name=sheet_name)
        except ValueError:
            # sheet不存在，创建新的sheet
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
                pd.DataFrame(columns=['']).to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"工作表不存在: 已在现有文件中创建了新工作表: {sheet_name}")

def rewrite(temp_dataframe, sheet):
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        temp_dataframe.to_excel(writer, sheet_name=sheet, index=False)
        # 将更新后的DataFrame写回到Excel文件


def append_row(sheet, **kwargs):           #表尾加一行
    ensure_exists(sheet)
    df = pd.read_excel(excel_file, sheet_name=sheet)
    new_row = pd.DataFrame([kwargs])
    df = pd.concat([df, new_row], ignore_index=True)
    
    rewrite(df, sheet)
    print(f"append: 已添加数据到工作表 {sheet}: {kwargs}")




def find_student(sheet, col, name):         #用任意一列的值查找
    ensure_exists(sheet)
    df = pd.read_excel(excel_file, sheet_name=sheet)
    if col not in df.columns:
        print(f"find: 列 '{col}' 在工作表 '{sheet}' 中不存在。")
        return None

    result = df[df[col] == name]
    if result.empty:
        print(f"find: 未找到学生 '{name}' 在工作表 '{sheet}' 中。")
        return None
    else:
        
        idp = df.index[df[col] == name].tolist()
        
        row_list = []
        for idx in idp:
            row_list.append(df.iloc[idx].to_dict())
            print(f"第{idx}行数据: {df.iloc[idx].to_dict()}")
        
        return pd.DataFrame(row_list, index=idp)


def delete_one_row(sheet, col, name):       #删除指定学生
    ensure_exists(sheet)
    df : DataFrame = pd.read_excel(excel_file, sheet_name=sheet)
    target_row : DataFrame = find_student(sheet, col, name)  # 提取符合条件的行
    if target_row is None:
        print(f"delete: 未找到学生 '{name}' 在工作表 '{sheet}' 中。")
        return None
    while True:
        try:
            indices_to_drop = list(map(int, input("delete: 输入希望删除的学生的行号索引（多个用空格分隔）: ").split()))
            if not all(idx in target_row.index for idx in indices_to_drop):
                print("delete: 输入的行号索引无效，请重新输入。")
                continue
            break
        except ValueError:
            print("delete: 输入的行号索引无效，请输入整数。")
            return None
    
    for idx in sorted(indices_to_drop, reverse=True):
            df.drop(index=idx, inplace=True)
            
        
    rewrite(df, sheet)
    if len(indices_to_drop) == 0:
        print(f"delete: 未删除任何行，因为输入的索引为空。")
    return len(indices_to_drop)



def delete_range(sheet, col, name):       #删除指定学生
    ensure_exists(sheet)
    df = pd.read_excel(excel_file, sheet_name=sheet)
    
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
        print(f"未找到学生 '{name}' 在工作表 '{sheet}' 中。")
        return 0
    
    rewrite(df, sheet)
    print(f"已删除学生 '{name}' 从工作表 '{sheet}'")
    return delete_length



# 测试
#append_row('students', Name='Tom', Score=90, Age=18)
#append_row('sheet1', Name='Jerry', Score=85)
delete_one_row('students', 'Name', 'Alex')
#print(delete_range('students', 'Name', 'Tom'))


'''
#输入部分
print("Welcome to the Score Manager!")
flag = 1
while flag:
    print("1. Add Score")
    print("2. View Scores")
    print("3. edit Score")
    print("4. Delete Score")
    print("5.find student")
    print("-1.exit") 
    choice = input("Enter your choice: ")
'''
'''
with pd.ExcelWriter('scores.xlsx') as writer:
    pd.DataFrame(columns=['Name', 'Score']).to_excel(writer, sheet_name='sheet1', index=False)
'''