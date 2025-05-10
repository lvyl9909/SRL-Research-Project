import pandas as pd
import numpy as np
from scipy import stats
import math
import argparse
import os

def compare_phase_proportions(file1, file2, phase_col='SRL_Phase', case_id_col='case_id'):
    """
    比较两个数据集中各Phase的比例，计算Mann-Whitney U检验的统计量
    
    参数:
    file1, file2 -- 两个数据集的Excel文件路径
    phase_col -- Phase列的名称(默认为'SRL_Phase')
    case_id_col -- 标识case的列名(默认为'case_id')
    """
    try:
        # 根据文件名自动确定哪个是stugptviz，哪个是recipe4u
        file1_name = os.path.basename(file1).lower()
        file2_name = os.path.basename(file2).lower()
        
        if 'stugptviz' in file1_name:
            stugptviz_file = file1
            recipe4u_file = file2
        elif 'stugptviz' in file2_name:
            stugptviz_file = file2
            recipe4u_file = file1
        else:
            print("警告: 无法根据文件名确定数据集类型，假设第一个是stugptviz，第二个是recipe4u")
            stugptviz_file = file1
            recipe4u_file = file2
        
        # 读取两个数据集
        print(f"读取Stugptviz数据集: {stugptviz_file}")
        stugptviz_df = pd.read_excel(stugptviz_file)
        
        print(f"读取Recipe4u数据集: {recipe4u_file}")
        recipe4u_df = pd.read_excel(recipe4u_file)
        
        # 检查必要的列是否存在
        required_columns = [phase_col, case_id_col]
        
        for col in required_columns:
            if col not in stugptviz_df.columns:
                print(f"错误: Stugptviz数据集中缺少 '{col}' 列")
                print(f"可用列: {', '.join(stugptviz_df.columns)}")
                return
        
        for col in required_columns:
            if col not in recipe4u_df.columns:
                print(f"错误: Recipe4u数据集中缺少 '{col}' 列")
                print(f"可用列: {', '.join(recipe4u_df.columns)}")
                return
        
        # 筛选有效的Phase值（去除缺失值）
        stugptviz_df = stugptviz_df.dropna(subset=[phase_col])
        recipe4u_df = recipe4u_df.dropna(subset=[phase_col])
        
        # 计算每个数据集中的唯一case_id数量
        stugptviz_n = stugptviz_df[case_id_col].nunique()
        recipe4u_n = recipe4u_df[case_id_col].nunique()
        
        print(f"Stugptviz组中唯一{case_id_col}数: {stugptviz_n}")
        print(f"Recipe4u组中唯一{case_id_col}数: {recipe4u_n}")
        
        # 确定所有唯一的Phase
        all_phases = sorted(set(stugptviz_df[phase_col].unique()) | set(recipe4u_df[phase_col].unique()))
        
        print(f"\n找到以下Phase类型: {', '.join(map(str, all_phases))}")
        
        # 创建结果数据框
        results = []
        
        # 计算每个case中各Phase的比例
        for phase in all_phases:
            # 对Stugptviz数据集，计算每个case中该Phase的比例
            stugptviz_ratios = []
            for case, group in stugptviz_df.groupby(case_id_col):
                total = len(group)
                if total > 0:
                    count = sum(group[phase_col] == phase)
                    ratio = (count / total) * 100
                    stugptviz_ratios.append(ratio)
            
            # 对Recipe4u数据集，计算每个case中该Phase的比例
            recipe4u_ratios = []
            for case, group in recipe4u_df.groupby(case_id_col):
                total = len(group)
                if total > 0:
                    count = sum(group[phase_col] == phase)
                    ratio = (count / total) * 100
                    recipe4u_ratios.append(ratio)
            
            # 如果两个数据集都有该Phase的数据
            if stugptviz_ratios and recipe4u_ratios:
                # 计算平均比例
                stugptviz_mean = np.mean(stugptviz_ratios)
                recipe4u_mean = np.mean(recipe4u_ratios)
                
                # 进行Mann-Whitney U检验
                u_stat, p_value = stats.mannwhitneyu(stugptviz_ratios, recipe4u_ratios, alternative='two-sided')
                
                # 计算平均秩
                combined = pd.DataFrame({
                    'ratio': stugptviz_ratios + recipe4u_ratios,
                    'group': ['Stugptviz'] * len(stugptviz_ratios) + ['Recipe4u'] * len(recipe4u_ratios)
                })
                combined['rank'] = combined['ratio'].rank()
                stugptviz_rank = combined[combined['group'] == 'Stugptviz']['rank'].mean()
                recipe4u_rank = combined[combined['group'] == 'Recipe4u']['rank'].mean()
                
                # 计算Z值
                n1 = len(stugptviz_ratios)
                n2 = len(recipe4u_ratios)
                z_value = (u_stat - (n1 * n2 / 2)) / math.sqrt(n1 * n2 * (n1 + n2 + 1) / 12)
                
                # 计算效应量 (r = Z / sqrt(N))
                effect_size = abs(z_value) / math.sqrt(n1 + n2)
                
                # 格式化p值，非常小的值显示为"<0.001"而不是0
                if p_value < 0.001:
                    p_value_str = "<0.001"
                elif p_value < 0.01:
                    p_value_str = f"{p_value:.3f}"
                elif p_value < 0.05:
                    p_value_str = f"{p_value:.3f}"
                else:
                    p_value_str = f"{p_value:.3f}"
                
                # 添加到结果
                results.append({
                    'Phase': phase,
                    'Mean Ratio (%) (Stugptviz, N = {})'.format(stugptviz_n): round(stugptviz_mean, 2),
                    'Mean Ratio (%) (Recipe4u, N = {})'.format(recipe4u_n): round(recipe4u_mean, 2),
                    'Mean Rank (Stugptviz, N = {})'.format(stugptviz_n): round(stugptviz_rank, 2),
                    'Mean Rank (Recipe4u, N = {})'.format(recipe4u_n): round(recipe4u_rank, 2),
                    'Z': round(z_value, 3),
                    'Effect Size (ES)': round(effect_size, 3),
                    'ZSig (2-tailed)': p_value_str  # 不再需要Significance列
                })
        
        # 将结果转换为DataFrame
        results_df = pd.DataFrame(results)
        
        # 添加显著性说明脚注，但不再需要星号说明
        # 如果您仍想保留脚注，可以改为只解释ZSig值
        footnote = "Note: ZSig values less than 0.001 are displayed as <0.001"
        
        # 显示结果
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', 1000)
        print("\n两组Phase比例的Mann-Whitney U检验结果:")
        print(results_df)
        print("\n" + footnote)
        
        # 保存到Excel文件
        output_file = "phase_comparison_results.xlsx"
        
        # 创建Excel写入器
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 写入主结果表
            results_df.to_excel(writer, sheet_name='Comparison Results', index=False)
            
            # 获取工作表
            worksheet = writer.sheets['Comparison Results']
            
            # 添加脚注
            footnote_row = len(results_df) + 3  # 留出2行空白
            worksheet.cell(row=footnote_row, column=1, value=footnote)
        
        print(f"\n结果已保存到: {output_file}")
        
        return results_df
        
    except Exception as e:
        print(f"处理过程中出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='比较两个数据集中SRL Phase的分布差异')
    parser.add_argument('file1', help='第一个Excel文件路径')
    parser.add_argument('file2', help='第二个Excel文件路径')
    parser.add_argument('-p', '--phase', default='SRL_Phase', help='Phase列的名称(默认为SRL_Phase)')
    parser.add_argument('-c', '--case', default='case_id', help='标识case的列名(默认为case_id)')
    
    args = parser.parse_args()
    compare_phase_proportions(args.file1, args.file2, args.phase, args.case)