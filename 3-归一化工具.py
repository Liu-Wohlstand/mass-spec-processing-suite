import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os

class NormalizeApp:
    def __init__(self, master):
        self.master = master
        master.title("质谱峰归一化工具（多分子列）")
        master.geometry("1100x700")

        self.last_dir = os.getcwd()
        self.df_current = None
        self.original_filename = ""

        # ---------- 文件选择 ----------
        self.file_frame = ttk.LabelFrame(master, text="选择文件", padding=5)
        self.file_frame.pack(fill="x", padx=5, pady=5)

        self.file_path = tk.StringVar()
        ttk.Entry(self.file_frame, textvariable=self.file_path, width=60).pack(side="left", padx=5)
        ttk.Button(self.file_frame, text="浏览", command=self.browse_file).pack(side="left")

        # ---------- 数据预览 ----------
        self.preview_frame = ttk.LabelFrame(master, text="数据预览（前20行）", padding=5)
        self.preview_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(self.preview_frame)
        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar = ttk.Scrollbar(self.preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # ---------- 列选择区域（三个列表 + 按钮） ----------
        self.col_frame = ttk.LabelFrame(master, text="分子 / 分母 选择", padding=5)
        self.col_frame.pack(fill="x", padx=5, pady=5)

        # 总列
        ttk.Label(self.col_frame, text="总列（所有列）").grid(row=0, column=0, padx=5, pady=5)
        self.all_listbox = tk.Listbox(self.col_frame, selectmode="multiple", height=8, width=25)
        self.all_listbox.grid(row=1, column=0, padx=5, pady=5)

        # 中间按钮
        btn_frame = ttk.Frame(self.col_frame)
        btn_frame.grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(btn_frame, text=">> 分子", command=self.add_to_mol).pack(pady=2)
        ttk.Button(btn_frame, text=">> 分母", command=self.add_to_den).pack(pady=2)
        ttk.Button(btn_frame, text="<< 移除分子", command=self.remove_from_mol).pack(pady=2)
        ttk.Button(btn_frame, text="<< 移除分母", command=self.remove_from_den).pack(pady=2)
        ttk.Button(btn_frame, text="清空分子", command=lambda: self.mol_listbox.delete(0, tk.END)).pack(pady=2)
        ttk.Button(btn_frame, text="清空分母", command=lambda: self.den_listbox.delete(0, tk.END)).pack(pady=2)

        # 分子列
        ttk.Label(self.col_frame, text="分子列（被归一化）").grid(row=0, column=2, padx=5, pady=5)
        self.mol_listbox = tk.Listbox(self.col_frame, selectmode="multiple", height=8, width=25)
        self.mol_listbox.grid(row=1, column=2, padx=5, pady=5)

        # 分母列
        ttk.Label(self.col_frame, text="分母列（求和）").grid(row=0, column=3, padx=5, pady=5)
        self.den_listbox = tk.Listbox(self.col_frame, selectmode="multiple", height=8, width=25)
        self.den_listbox.grid(row=1, column=3, padx=5, pady=5)

        # ---------- 操作按钮 ----------
        self.button_frame = ttk.Frame(master)
        self.button_frame.pack(fill="x", padx=5, pady=5)

        ttk.Button(self.button_frame, text="执行归一化", command=self.normalize).pack(side="left", padx=5)
        ttk.Button(self.button_frame, text="保存结果", command=self.save_file).pack(side="left", padx=5)

    # ---------- 文件浏览 ----------
    def browse_file(self):
        filename = filedialog.askopenfilename(
            initialdir=self.last_dir,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filename:
            self.file_path.set(filename)
            self.last_dir = os.path.dirname(filename)
            self.original_filename = filename
            self.load_file(filename)

    # ---------- 加载文件 ----------
    def load_file(self, filename):
        try:
            self.df_current = pd.read_excel(filename, engine="openpyxl")
            self.update_preview(self.df_current)
            self.update_all_columns()
            # 清空分子/分母列表
            self.mol_listbox.delete(0, tk.END)
            self.den_listbox.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件：{e}")

    # ---------- 更新预览 ----------
    def update_preview(self, df, max_rows=20):
        for row in self.tree.get_children():
            self.tree.delete(row)
        cols = list(df.columns)
        self.tree["columns"] = cols
        self.tree["show"] = "headings"
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
        for i, row in df.head(max_rows).iterrows():
            values = [str(v) for v in row.values]
            self.tree.insert("", "end", values=values)
        if len(df) > max_rows:
            self.tree.insert("", "end", values=["... 共 {} 行，仅显示前 {} 行 ...".format(len(df), max_rows)])

    # ---------- 更新总列列表 ----------
    def update_all_columns(self):
        self.all_listbox.delete(0, tk.END)
        if self.df_current is not None:
            for col in self.df_current.columns:
                self.all_listbox.insert(tk.END, col)

    # ---------- 添加选中列到分子 ----------
    def add_to_mol(self):
        selected = self.all_listbox.curselection()
        if not selected:
            return
        current_mol = list(self.mol_listbox.get(0, tk.END))
        for idx in selected:
            col = self.all_listbox.get(idx)
            if col not in current_mol:
                self.mol_listbox.insert(tk.END, col)
        # 可选：清空总列选择
        self.all_listbox.selection_clear(0, tk.END)

    # ---------- 添加选中列到分母 ----------
    def add_to_den(self):
        selected = self.all_listbox.curselection()
        if not selected:
            return
        current_den = list(self.den_listbox.get(0, tk.END))
        for idx in selected:
            col = self.all_listbox.get(idx)
            if col not in current_den:
                self.den_listbox.insert(tk.END, col)
        self.all_listbox.selection_clear(0, tk.END)

    # ---------- 从分子移除 ----------
    def remove_from_mol(self):
        selected = self.mol_listbox.curselection()
        if not selected:
            return
        # 从后往前删除，避免索引变化
        for idx in reversed(selected):
            self.mol_listbox.delete(idx)

    # ---------- 从分母移除 ----------
    def remove_from_den(self):
        selected = self.den_listbox.curselection()
        if not selected:
            return
        for idx in reversed(selected):
            self.den_listbox.delete(idx)

    # ---------- 归一化核心 ----------
    def normalize(self):
        if self.df_current is None:
            messagebox.showwarning("警告", "请先加载文件！")
            return

        mol_cols = list(self.mol_listbox.get(0, tk.END))
        den_cols = list(self.den_listbox.get(0, tk.END))

        if not mol_cols:
            messagebox.showwarning("警告", "分子列为空！请先添加分子列。")
            return
        if not den_cols:
            messagebox.showwarning("警告", "分母列为空！请先添加分母列。")
            return

        # 检查是否存在重叠（分母包含分子列会导致除零，建议警告但允许）
        overlap = set(mol_cols) & set(den_cols)
        if overlap:
            if not messagebox.askyesno("警告", f"下列列同时出现在分子和分母中：{overlap}\n这可能导致分母包含自身，确定继续吗？"):
                return

        # 检查列是否存在
        current_cols = list(self.df_current.columns)
        for col in mol_cols + den_cols:
            if col not in current_cols:
                messagebox.showerror("错误", f"列 '{col}' 不存在于当前数据中！")
                return

        # 筛选出尚未归一化的分子列（避免重复插入）
        to_process = []
        for mol_col in mol_cols:
            norm_col = mol_col + "_归一化"
            if norm_col in current_cols:
                messagebox.showinfo("提示", f"列 '{norm_col}' 已存在，将跳过分子列 '{mol_col}'。")
            else:
                to_process.append(mol_col)

        if not to_process:
            messagebox.showinfo("提示", "所有选中的分子列都已归一化，无需操作。")
            return

        # 计算分母之和
        try:
            denominator = self.df_current[den_cols].sum(axis=1)
            denominator = denominator.replace(0, np.nan)
        except Exception as e:
            messagebox.showerror("错误", f"分母计算失败：{e}")
            return

        # 构建新列数据
        new_columns = {}
        for mol_col in to_process:
            try:
                numerator = self.df_current[mol_col]
                norm_val = numerator / denominator
                norm_val = norm_val.fillna(0).replace([np.inf, -np.inf], 0)
                new_columns[mol_col] = norm_val
            except Exception as e:
                messagebox.showerror("错误", f"计算分子列 '{mol_col}' 失败：{e}")
                return

        # 重新构建列顺序：在每个分子列后面插入归一化列
        old_cols = list(self.df_current.columns)
        new_col_order = []
        for col in old_cols:
            new_col_order.append(col)
            if col in to_process:
                new_col_order.append(col + "_归一化")

        # 创建新DataFrame
        df_new = pd.DataFrame()
        for col in new_col_order:
            if col.endswith("_归一化"):
                base_col = col[:-4]
                if base_col in new_columns:
                    df_new[col] = new_columns[base_col]
                else:
                    df_new[col] = np.nan
            else:
                df_new[col] = self.df_current[col]

        self.df_current = df_new
        self.update_preview(self.df_current)
        self.update_all_columns()   # 刷新总列列表（因为列名改变了）
        # 注意：分子和分母列表中的列名可能包含已存在的归一化列，但不会自动更新，我们保持现有选择不变（用户可手动调整）
        messagebox.showinfo("完成", f"成功为 {len(to_process)} 个分子列添加了归一化列。")

    # ---------- 保存文件 ----------
    def save_file(self):
        if self.df_current is None:
            messagebox.showwarning("警告", "请先加载并执行归一化！")
            return
        if not self.original_filename:
            return
        base, ext = os.path.splitext(self.original_filename)
        new_path = base + "_归一化" + ext
        try:
            self.df_current.to_excel(new_path, index=False, engine="openpyxl")
            messagebox.showinfo("保存成功", f"文件已保存到：\n{new_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = NormalizeApp(root)
    root.mainloop()