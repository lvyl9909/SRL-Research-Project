import pandas as pd
import numpy as np
import argparse
import os

def analyze_transition_matrix(file_path, case_id_col='case_id', action_col='SRL_code'):
    """
    分析Excel数据集中的动作转换情况，创建转换矩阵并显示频率和百分比
    
    参数:
    file_path -- Excel文件路径
    case_id_col -- 标识case的列名(默认为'case_id')
    action_col -- 动作代码的列名(默认为'SRL_code')
    """
    try:
        # 读取Excel文件
        print(f"读取文件: {file_path}")
        df = pd.read_excel(file_path)
        
        # 检查必要的列是否存在
        required_columns = [case_id_col, action_col]
        for col in required_columns:
            if col not in df.columns:
                print(f"错误: 列 '{col}' 不存在于Excel文件中。")
                print(f"可用列: {', '.join(df.columns)}")
                return
        
        # 过滤掉R.SL *，而不是转换它
        print(f"过滤前行数: {len(df)}")
        df = df[df[action_col] != 'R.SL *']
        print(f"过滤后行数: {len(df)}")
        
        # 定义代码顺序 - 按照指定的排序
        code_order = [
            "F.DP", "F.SG", 
            "M.CU(B)", "M.CU(D)", 
            "C.SC(B)", "C.SC(D)", 
            "C.RH(I)", "C.RH(E)", "C.RH(C)",
            "C.AI", "C.RP", "C.CA", 
            "R.SE", "R.SL", "R.PN"
        ]
        
        # 检查数据集中包含的代码
        existing_codes = set(df[action_col].unique())
        valid_codes = [code for code in code_order if code in existing_codes]
        
        # 如果有数据集中出现但不在预定义顺序中的代码，将它们添加到末尾
        missing_codes = [code for code in existing_codes if code not in code_order]
        if missing_codes:
            print(f"警告: 数据集中有未在预定义顺序中的代码: {', '.join(missing_codes)}")
            valid_codes.extend(missing_codes)
        
        # 创建一个空的转换矩阵
        transition_matrix = pd.DataFrame(0, index=valid_codes, columns=valid_codes)
        
        # 计算每个case内的转换
        for case_id, group in df.groupby(case_id_col):
            # 确保按照原始顺序排序
            group = group.sort_index()
            
            # 获取动作序列
            actions = group[action_col].values
            
            # 计算转换
            for i in range(len(actions) - 1):
                current_action = actions[i]
                next_action = actions[i + 1]
                
                # 确保动作在我们的矩阵中
                if current_action in valid_codes and next_action in valid_codes:
                    transition_matrix.loc[current_action, next_action] += 1
        
        # 重新按照指定顺序排列代码
        # 只使用数据集中实际存在的代码，并保持指定的顺序
        ordered_codes = [code for code in code_order if code in valid_codes]
        
        # 重新排序转换矩阵
        transition_matrix = transition_matrix.reindex(index=ordered_codes, columns=ordered_codes)
        
        # 计算每列的和（作为后续行为出现的次数）
        column_sums = transition_matrix.sum(axis=0)
        total_transitions = column_sums.sum()
        
        # 计算每列的百分比
        percentage_cols = (column_sums / total_transitions * 100).round(2)
        
        # 创建带有频率和百分比行的扩展矩阵
        extended_matrix = pd.DataFrame(index=['N', '%'] + ordered_codes, columns=ordered_codes)
        
        # 填充频率和百分比行（基于列和）
        for code in ordered_codes:
            extended_matrix.loc['N', code] = column_sums[code]
            extended_matrix.loc['%', code] = f"{percentage_cols[code]}%"
        
        # 填充转换数据
        for row_code in ordered_codes:
            for col_code in ordered_codes:
                extended_matrix.loc[row_code, col_code] = transition_matrix.loc[row_code, col_code]
        
        # 保存到Excel文件
        output_file = os.path.basename(file_path).replace('.xlsx', '_column_transition_matrix.xlsx')
        extended_matrix.to_excel(output_file)
        
        print(f"转换矩阵已保存到: {output_file}")
        print(f"总转换次数: {total_transitions}")
        print(f"唯一代码数: {len(valid_codes)}")
        
        return extended_matrix
        
    except Exception as e:
        print(f"处理过程中出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='分析Excel数据集中的动作转换情况')
    parser.add_argument('file', help='Excel文件路径')
    parser.add_argument('-c', '--case', default='case_id', help='标识case的列名(默认为case_id)')
    parser.add_argument('-a', '--action', default='SRL_code', help='动作代码的列名(默认为SRL_code)')
    
    args = parser.parse_args()
    analyze_transition_matrix(args.file, args.case, args.action)