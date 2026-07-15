import matplotlib
matplotlib.use('TkAgg')

import numpy as np
import pandas as pd
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt
import warnings
import os
import tkinter as tk
from tkinter import filedialog
warnings.filterwarnings('ignore')

# ========== 选择文件 ==========
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(
    title="Select Excel file (.xlsx)",
    filetypes=[("Excel files", "*.xlsx")]
)
if not file_path:
    print("No file selected.")
    exit()

# ========== 读取数据 ==========
df_data = pd.read_excel(file_path, header=0)
col_x, col_y = df_data.columns[0], df_data.columns[1]
x = df_data[col_x].values
y = df_data[col_y].values
n = len(x)

# ========== 定义拟合函数 ==========
functions = [
    {'name': 'Linear', 'func': lambda x, a, b: a*x + b, 'p0': [0,0]},
    {'name': 'Quadratic', 'func': lambda x, a, b, c: a*x**2 + b*x + c, 'p0': [0,0,0]},
    {'name': 'Cubic', 'func': lambda x, a, b, c, d: a*x**3 + b*x**2 + c*x + d, 'p0': [0,0,0,0]},
    {'name': 'Quartic', 'func': lambda x, a, b, c, d, e: a*x**4 + b*x**3 + c*x**2 + d*x + e, 'p0': [0,0,0,0,0]},
    {'name': 'Quintic', 'func': lambda x, a, b, c, d, e, f: a*x**5 + b*x**4 + c*x**3 + d*x**2 + e*x + f, 'p0': [0,0,0,0,0,0]},
    {'name': 'Exponential', 'func': lambda x, a, b: a * np.exp(b*x), 'p0': [0.005, -0.01]},
    {'name': 'ExpConst', 'func': lambda x, a, b, c: a * np.exp(b*x) + c, 'p0': [0.005, -0.01, 0]},
    {'name': 'DoubleExp', 'func': lambda x, a, b, c, d: a*np.exp(b*x) + c*np.exp(d*x), 'p0': [0.005, -0.01, 0.005, -0.02]},
    {'name': 'Power', 'func': lambda x, a, b: a * x**b, 'p0': [0.01, -0.5]},
    {'name': 'Logarithmic', 'func': lambda x, a, b: a * np.log(x) + b, 'p0': [-0.001, 0.01]},
    {'name': 'SquareRoot', 'func': lambda x, a, b: a * np.sqrt(x) + b, 'p0': [-0.001, 0.01]},
    {'name': 'Reciprocal', 'func': lambda x, a, b: a / x + b, 'p0': [0.1, 0]},
    {'name': 'Logistic', 'func': lambda x, a, b, c: a / (1 + np.exp(-b*(x-c))), 'p0': [0.005, -0.1, 50]},
    {'name': 'Gaussian', 'func': lambda x, a, b, c: a * np.exp(-(x-b)**2/(2*c**2)), 'p0': [0.005, 50, 20]},
    {'name': 'Rational', 'func': lambda x, a, b, c, d: (a*x + b) / (c*x + d), 'p0': [0, 0.005, 0, 1]}
]

# ========== 拟合与评估 ==========
results = []
ss_tot = np.sum((y - np.mean(y))**2)
for item in functions:
    try:
        popt, _ = curve_fit(item['func'], x, y, p0=item['p0'], maxfev=5000)
        y_pred = item['func'](x, *popt)
        ss_res = np.sum((y - y_pred)**2)
        r2 = 1 - ss_res/ss_tot
        k = len(popt)
        r2_adj = 1 - (1-r2)*(n-1)/(n-k-1) if n-k-1 > 0 else -np.inf
        results.append({
            'Model': item['name'],
            'k': k,
            'R2': r2,
            'Adj_R2': r2_adj,
            'params': popt,
            'func_obj': item['func']
        })
    except:
        pass

df = pd.DataFrame(results)
df_sorted = df.sort_values('Adj_R2', ascending=False).reset_index(drop=True)

# ========== 辅助函数：生成公式字符串 ==========
def format_equation(name, params):
    """返回可读的公式字符串，例如 'y = -0.0008*sqrt(x) + 0.0097'"""
    p = params
    if name == 'Linear':
        return f"y = {p[0]:.4f}*x + {p[1]:.4f}"
    elif name == 'Quadratic':
        return f"y = {p[0]:.4f}*x² + {p[1]:.4f}*x + {p[2]:.4f}"
    elif name == 'Cubic':
        return f"y = {p[0]:.4f}*x³ + {p[1]:.4f}*x² + {p[2]:.4f}*x + {p[3]:.4f}"
    elif name == 'Quartic':
        return f"y = {p[0]:.4f}*x⁴ + {p[1]:.4f}*x³ + {p[2]:.4f}*x² + {p[3]:.4f}*x + {p[4]:.4f}"
    elif name == 'Quintic':
        return f"y = {p[0]:.4f}*x⁵ + {p[1]:.4f}*x⁴ + {p[2]:.4f}*x³ + {p[3]:.4f}*x² + {p[4]:.4f}*x + {p[5]:.4f}"
    elif name == 'Exponential':
        return f"y = {p[0]:.4f}*exp({p[1]:.4f}*x)"
    elif name == 'ExpConst':
        return f"y = {p[0]:.4f}*exp({p[1]:.4f}*x) + {p[2]:.4f}"
    elif name == 'DoubleExp':
        return f"y = {p[0]:.4f}*exp({p[1]:.4f}x) + {p[2]:.4f}*exp({p[3]:.4f}x)"
    elif name == 'Power':
        return f"y = {p[0]:.4f}*x^{p[1]:.4f}"
    elif name == 'Logarithmic':
        return f"y = {p[0]:.4f}*ln(x) + {p[1]:.4f}"
    elif name == 'SquareRoot':
        return f"y = {p[0]:.6f}*sqrt(x) + {p[1]:.6f}"
    elif name == 'Reciprocal':
        return f"y = {p[0]:.4f}/x + {p[1]:.4f}"
    elif name == 'Logistic':
        return f"y = {p[0]:.4f}/(1+exp(-{p[1]:.4f}*(x-{p[2]:.4f})))"
    elif name == 'Gaussian':
        return f"y = {p[0]:.4f}*exp(-(x-{p[1]:.4f})²/(2*{p[2]:.4f}²))"
    elif name == 'Rational':
        return f"y = ({p[0]:.4f}*x+{p[1]:.4f})/({p[2]:.4f}*x+{p[3]:.4f})"
    else:
        return ""

# ========== 输出到控制台和txt ==========
output_lines = []
output_lines.append(f"Data: {os.path.basename(file_path)}, x='{col_x}', y='{col_y}', n={n}\n")
output_lines.append("=== Fitting Results (sorted by Adjusted R²) ===")
output_lines.append(f"{'Model':<15} {'k':<3} {'R²':<10} {'Adj R²':<10}")
output_lines.append("-"*50)
for _, row in df_sorted.iterrows():
    output_lines.append(f"{row['Model']:<15} {row['k']:<3} {row['R2']:<10.6f} {row['Adj_R2']:<10.6f}")

best = df_sorted.iloc[0]
output_lines.append(f"\n✅ Best model: {best['Model']} (Adj R²={best['Adj_R2']:.6f})")
output_lines.append(f"   Equation: {format_equation(best['Model'], best['params'])}")

# 保存txt
base_dir = os.path.dirname(file_path)
base_name = os.path.splitext(os.path.basename(file_path))[0]
txt_path = os.path.join(base_dir, base_name + "_fitting_results.txt")
with open(txt_path, 'w', encoding='utf-8') as f:
    f.write("\n".join(output_lines))
print("\n".join(output_lines))
print(f"\nResults saved to: {txt_path}")

# ========== 绘图（图例显示公式和R²）==========
plt.figure(figsize=(8,6))
plt.scatter(x, y, label='Original Data', color='black', s=50, zorder=5)

x_smooth = np.linspace(min(x)-1, max(x)+1, 300)
top_n = 1  # 显示前1个模型
for i in range(min(top_n, len(df_sorted))):
    row = df_sorted.iloc[i]
    name = row['Model']
    params = row['params']
    func = row['func_obj']
    y_smooth = func(x_smooth, *params)
    adj_r2 = row['Adj_R2']
    # 构造图例标签：模型名 + 公式 + R²
    eq_str = format_equation(name, params)
    label = f"{name}: {eq_str}\nAdj R² = {adj_r2:.4f}"
    plt.plot(x_smooth, y_smooth, linewidth=2, label=label)

plt.xlabel(col_x, fontsize=12)
plt.ylabel(col_y, fontsize=12)
plt.title('Fitting Curves Comparison (based on Adjusted R²)', fontsize=14)
plt.legend(loc='best', fontsize=8)  # 字体调小以防遮挡
plt.grid(alpha=0.3)
plt.tight_layout()

fig_path = os.path.join(base_dir, base_name + "_fitting_plot.png")
plt.savefig(fig_path, dpi=300)
print(f"Plot saved to: {fig_path}")
plt.show()