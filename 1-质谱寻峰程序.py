import pandas as pd
import numpy as np
import os
import glob
import tkinter as tk
from tkinter import filedialog

# ==================== 用户配置区 ====================
# 在此处设置需要检索的质荷比（m/z）列表，支持整数或小数
target_mz_list = [126.90502, 144.91975, 189.90067]   # 示例值，请根据实际需要修改

# 搜索窗口半径（质荷比单位），根据仪器分辨率调整，此处设为 ±0.01
delta_mz = 0.05

# 记忆文件名称（保存在脚本同目录下）
MEMORY_FILE = "searchpeak_last_folder.txt"
# ===================================================

def get_folder_path():
    """
    弹出文件夹选择对话框，返回用户选择的路径。
    自动记忆上次选择的路径，下次打开时默认定位到该目录。
    若用户取消选择，则返回当前工作目录。
    """
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 读取上次保存的路径
    last_path = ""
    if os.path.exists(MEMORY_FILE):
        try:
            with open(MEMORY_FILE, 'r') as f:
                last_path = f.read().strip()
        except:
            pass

    # 若上次路径无效，则使用当前目录
    if not os.path.isdir(last_path):
        last_path = os.getcwd()

    # 弹出文件夹选择对话框，初始目录设为上次路径
    folder_selected = filedialog.askdirectory(
        title="选择包含质谱数据的文件夹",
        initialdir=last_path
    )

    if folder_selected:  # 用户选择了有效路径
        # 保存这次选择的路径以便下次使用
        with open(MEMORY_FILE, 'w') as f:
            f.write(folder_selected)
        return folder_selected
    else:
        # 用户取消，返回当前目录
        return os.getcwd()


def find_peak_intensity(mz_array, intensity_array, target_mz, delta):
    """
    在给定的质荷比数组中，搜索 target_mz 附近 ±delta 范围内的最高强度。
    若范围内无数据点，则返回 0.0。
    """
    left_idx = np.searchsorted(mz_array, target_mz - delta)
    right_idx = np.searchsorted(mz_array, target_mz + delta)
    if left_idx < right_idx:
        return np.max(intensity_array[left_idx:right_idx])
    else:
        return 0.0


def process_file(file_path, target_mz_list, delta):
    """读取单个 Excel 文件，返回每个目标 m/z 对应的最高峰强度列表"""
    try:
        df = pd.read_excel(file_path, header=0)
        # 根据数据描述，第二列为质荷比，第三列为强度
        mz_array = df.iloc[:, 1].values.astype(float)
        intensity_array = df.iloc[:, 2].values.astype(float)
    except Exception as e:
        print(f"读取文件 {file_path} 时出错: {e}")
        return [None] * len(target_mz_list)

    results = []
    for mz in target_mz_list:
        peak_intensity = find_peak_intensity(mz_array, intensity_array, mz, delta)
        results.append(peak_intensity)
    return results


def main():
    # 获取用户选择的文件夹（带记忆功能）
    folder_path = get_folder_path()
    print(f"工作文件夹: {folder_path}\n")

    # 获取所有 .xlsx 文件
    pattern = os.path.join(folder_path, "*.xlsx")
    file_list = glob.glob(pattern)

    if not file_list:
        print(f"在文件夹 {folder_path} 中未找到任何 .xlsx 文件")
        return

    print(f"找到 {len(file_list)} 个 Excel 文件\n")
    print("=" * 80)

    summary_data = []
    headers = ["文件名"] + [f"m/z {mz}" for mz in target_mz_list]

    for file_path in file_list:
        file_name = os.path.basename(file_path)
        print(f"处理文件: {file_name}")

        peaks = process_file(file_path, target_mz_list, delta_mz)
        # 打印该文件的结果
        print(f"  结果: {dict(zip(target_mz_list, peaks))}")

        row = [file_name] + peaks
        summary_data.append(row)

    # 创建汇总 DataFrame
    df_summary = pd.DataFrame(summary_data, columns=headers)

    # 保存汇总 Excel 文件到当前工作目录（或文件夹选择目录，可根据需要调整）
    output_file = os.path.join(folder_path, "peak_intensity_summary.xlsx")
    df_summary.to_excel(output_file, index=False)

    print("\n" + "=" * 80)
    print(f"汇总结果已保存至: {output_file}")

    # 打印汇总表格预览
    print("\n汇总表格预览:")
    print(df_summary.to_string(index=False))


if __name__ == "__main__":
    main()
