import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.signal import savgol_filter
from rdp import rdp
import os
import re

# ==============================================================================
# --- 1. 配置区 (请在此处修改) ---
# ==============================================================================
# --- 输入文件 ---
SOURCE_FILENAME = r"C:\Users\asus\Desktop\AAA-University\科研\人工肌肉项目\数据\连续测量\PET+乳胶管，215mm.xlsx"
# --- 算法参数 ---
# 预处理分箱数量：解决“回头线”和大部分噪声的关键。推荐 50 ~ 200。
PREPROCESS_BINS = 100

# Savitzky-Golay 滤波器参数
SG_WINDOW_LENGTH = 28 # 窗口应为奇数，且小于 PREPROCESS_BINS
SG_POLYORDER = 3
# 自适应 RDP 算法的比例系数
# 值越小，保留的点越多；值越大，保留的点越少。
RDP_ADAPTIVE_RATIO = 0.0005

# --- 图表配置 ---
CHART_TITLE = "PET+Latex Tube"
X_AXIS_LABEL = "Force (N)"
Y_AXIS_LABEL = "Contraction Rate (%)"
LEGEND_TITLE = "Pressure(MPa)"
COLOR_PALETTE = 'mako' # 推荐: 'viridis', 'plasma', 'mako'

# --- 字体设置 ---
try:
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
except:
    print("警告：未找到 'SimHei' 字体，中文可能无法正常显示。")

# ==============================================================================
# --- 2. 核心功能函数 (无需修改) ---
# ==============================================================================

def preprocess_data(x, y, num_bins):
    """
    对原始数据进行“排序-分组-平均”预处理，消除噪声和数据分支。
    """
    if len(x) < 2: return x, y
    points = sorted(zip(x, y))
    x_sorted, y_sorted = np.array(list(zip(*points)))
    bins = np.linspace(x_sorted[0], x_sorted[-1], num_bins + 1)
    digitized = np.digitize(x_sorted, bins)
    x_binned = [x_sorted[digitized == i].mean() for i in range(1, len(bins))]
    y_binned = [y_sorted[digitized == i].mean() for i in range(1, len(bins))]
    x_processed = np.array([val for val in x_binned if not np.isnan(val)])
    y_processed = np.array([val for val in y_binned if not np.isnan(val)])
    return x_processed, y_processed

def parse_data_blocks(df):
    """
    从DataFrame中动态解析出所有数据块。
    """
    data_blocks = {}
    for c in range(df.shape[1]):
        pressure_cell = df[c].astype(str).str.contains('气压')
        if pressure_cell.any():
            row_idx, col_idx = pressure_cell.idxmax(), c
            pressure_title = str(df.iat[row_idx, col_idx])
            
            # 使用正则表达式从 '气压: 0.1 MPa' 中提取出 '0.1'
            match = re.search(r'[\d\.]+', pressure_title)
            clean_label = match.group(0) if match else f"系列{c+1}"

            print(f"\n找到数据块: '{pressure_title}' -> 解析为标签: '{clean_label}'")
            data_start_row = row_idx + 2
            force_data = pd.to_numeric(df.iloc[data_start_row:, col_idx], errors='coerce').dropna()
            shrinkage_data = pd.to_numeric(df.iloc[data_start_row:, col_idx + 1], errors='coerce').dropna()
            
            if not force_data.empty and not shrinkage_data.empty:
                data_blocks[clean_label] = (force_data.values, shrinkage_data.values)
                print(f"解析成功，找到 {len(force_data)} 个有效数据点。")
    return data_blocks

# ==============================================================================
# --- 3. 主处理与绘图流程 ---
# ==============================================================================

def main_process_and_plot(source_file):
    """
    加载、处理并统一绘制所有数据系列。
    """
    # --- 步骤 1: 加载并解析数据 ---
    try:
        raw_df = pd.read_excel(source_file, header=None)
        print(f"成功加载文件: '{source_file}'")
    except FileNotFoundError:
        print(f"错误: 文件 '{source_file}' 未找到。")
        return
        
    data_blocks = parse_data_blocks(raw_df)
    if not data_blocks:
        print("\n处理错误：在文件中未能解析出任何有效的数据块。")
        return

    # --- 步骤 2: 创建统一的对比图表 ---
    fig, ax = plt.subplots(figsize=(12, 8))
    sns.set_theme(style='whitegrid')
    colors = sns.color_palette(COLOR_PALETTE, n_colors=len(data_blocks))

    # --- 步骤 3: 循环处理每个数据块并绘图 ---
    for i, (label, (force_raw, shrinkage_raw)) in enumerate(data_blocks.items()):
        print(f"\n--- 正在处理: {label} MPa ---")
        
        # --- 核心升级 1: 数据预处理 (重排) ---
        print(f"应用'排序-分组-平均'预处理 (分成 {PREPROCESS_BINS} 组)...")
        force_proc, shrinkage_proc = preprocess_data(force_raw, shrinkage_raw, num_bins=PREPROCESS_BINS)
        
        if len(force_proc) < SG_WINDOW_LENGTH:
            print(f"警告：预处理后数据点 ({len(force_proc)}) 少于SG滤波器窗口 ({SG_WINDOW_LENGTH})，跳过此压力档。")
            continue
        
        # --- 核心升级 2: 在预处理后的数据上进行平滑 ---
        print(f"应用Savitzky-Golay滤波器...")
        force_smooth = savgol_filter(force_proc, SG_WINDOW_LENGTH, SG_POLYORDER)
        shrinkage_smooth = savgol_filter(shrinkage_proc, SG_WINDOW_LENGTH, SG_POLYORDER)
        
        # --- 保留的优点: 自适应 RDP ---
        shrinkage_range = np.max(shrinkage_smooth) - np.min(shrinkage_smooth)
        dynamic_epsilon = RDP_ADAPTIVE_RATIO * shrinkage_range + 1e-9
        print(f"动态计算 Epsilon: {dynamic_epsilon:.5f}")
        
        smooth_points = np.column_stack([force_smooth, shrinkage_smooth])
        key_points = rdp(smooth_points, epsilon=dynamic_epsilon)
        
        force_final, shrinkage_final = key_points.T
        print(f"处理完成！关键点数: {len(force_final)}")
        
        # --- 核心升级 3: 在同一个图表上绘制最终结果 ---
        ax.plot(force_final, shrinkage_final, 'o-', 
                label=label, 
                color=colors[i], 
                linewidth=2, 
                markersize=5)

    # --- 步骤 4: 统一美化最终的图表 ---
    ax.set_title(CHART_TITLE, fontsize=18, weight='bold', pad=20)
    ax.set_xlabel(X_AXIS_LABEL, fontsize=14)
    ax.set_ylabel(Y_AXIS_LABEL, fontsize=14)
    ax.legend(title=LEGEND_TITLE, fontsize=11, title_fontsize=13)
    ax.tick_params(axis='both', which='major', labelsize=12)
    ax.grid(True, which='major', linestyle='--', linewidth=0.5)
    plt.tight_layout()
    plt.show()

# ==============================================================================
# --- 4. 程序入口 ---
# ==============================================================================
if __name__ == "__main__":
    main_process_and_plot(SOURCE_FILENAME)