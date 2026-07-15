import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
import numpy as np
import matplotlib

matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import json
import os

# ===== 设置中文字体 =====
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'WenQuanYi Zen Hei']
plt.rcParams['axes.unicode_minus'] = False
# ========================
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

CONFIG_FILE = "energy_profile_config.json"


class EnergyProfileApp:
    def __init__(self, master):
        self.master = master
        master.title("能垒图绘制工具（多文件）")
        master.geometry("1300x850")

        # ---------- 加载保存的目录 ----------
        self.last_dir = os.getcwd()
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    if 'last_dir' in config and os.path.exists(config['last_dir']):
                        self.last_dir = config['last_dir']
            except:
                pass

        self.files = {}
        self.file_names = []
        self.current_file = None
        self.curves = []
        self.auto_color_idx = 0
        self.colors = ['blue', 'red', 'green', 'orange', 'purple', 'brown', 'pink', 'gray', 'olive', 'cyan']

        # ---------- 文件管理 ----------
        file_frame = ttk.LabelFrame(master, text="文件管理", padding=5)
        file_frame.pack(fill="x", padx=5, pady=5)

        self.file_listbox = tk.Listbox(file_frame, height=4)
        self.file_listbox.pack(side="left", fill="x", expand=True, padx=5)
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)

        btn_file_frame = ttk.Frame(file_frame)
        btn_file_frame.pack(side="left", padx=5)
        ttk.Button(btn_file_frame, text="添加文件", command=self.add_file).pack(pady=2)
        ttk.Button(btn_file_frame, text="移除文件", command=self.remove_file).pack(pady=2)

        # ---------- 数据预览 ----------
        preview_frame = ttk.LabelFrame(master, text="数据预览（当前文件，前20行）", padding=5)
        preview_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(preview_frame)
        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # ---------- 曲线管理 ----------
        curve_frame = ttk.LabelFrame(master, text="曲线管理（双击编辑）", padding=5)
        curve_frame.pack(fill="x", padx=5, pady=5)

        self.curve_listbox = tk.Listbox(curve_frame, height=4)
        self.curve_listbox.pack(side="left", fill="x", expand=True, padx=5)
        self.curve_listbox.bind('<Double-Button-1>', self.edit_curve)

        btn_curve_frame = ttk.Frame(curve_frame)
        btn_curve_frame.pack(side="left", padx=5)
        ttk.Button(btn_curve_frame, text="添加曲线", command=self.add_curve).pack(pady=2)
        ttk.Button(btn_curve_frame, text="删除曲线", command=self.delete_curve).pack(pady=2)
        ttk.Button(btn_curve_frame, text="清空所有", command=self.clear_curves).pack(pady=2)

        # ---------- 绘图参数 ----------
        param_frame = ttk.LabelFrame(master, text="绘图参数", padding=5)
        param_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(param_frame, text="偏移量 m (水平半宽):").pack(side="left", padx=5)
        self.m_var = tk.DoubleVar(value=0.5)
        ttk.Entry(param_frame, textvariable=self.m_var, width=10).pack(side="left", padx=5)

        ttk.Button(param_frame, text="绘制预览", command=self.plot_preview).pack(side="left", padx=20)
        ttk.Button(param_frame, text="保存图片（弹出窗口）", command=self.save_figure).pack(side="left", padx=5)

        # ---------- 绘图画布（仅预览） ----------
        fig_frame = ttk.LabelFrame(master, text="绘图预览", padding=5)
        fig_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.fig, self.ax = plt.subplots(figsize=(6, 4))
        self.fig.set_tight_layout(True)
        self.canvas = FigureCanvasTkAgg(self.fig, master=fig_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # ---------- Sizegrip ----------
        self.sizegrip = ttk.Sizegrip(master)
        self.sizegrip.pack(side="bottom", anchor="se")

        self.df = None

        # ---------- 窗口关闭时保存配置 ----------
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ---------- 持久化存储 ----------
    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump({'last_dir': self.last_dir}, f)
        except:
            pass

    def on_closing(self):
        self.save_config()
        self.master.destroy()

    # ---------- 文件操作 ----------
    def add_file(self):
        filename = filedialog.askopenfilename(
            initialdir=self.last_dir,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not filename:
            return
        # 更新并保存目录
        self.last_dir = os.path.dirname(filename)
        self.save_config()

        if filename in self.files:
            messagebox.showinfo("提示", "该文件已加载！")
            return
        try:
            df = pd.read_excel(filename, engine="openpyxl")
            self.files[filename] = df
            self.file_names.append(filename)
            self.file_listbox.insert(tk.END, os.path.basename(filename))
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(tk.END)
            self.on_file_select()
            messagebox.showinfo("成功", f"加载文件：{os.path.basename(filename)}，共 {len(df)} 行。")
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败：{e}")

    def remove_file(self):
        selected = self.file_listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        filename = self.file_names[idx]
        del self.files[filename]
        del self.file_names[idx]
        self.file_listbox.delete(idx)
        self.curves = [c for c in self.curves if c['file'] != filename]
        self.curve_listbox.delete(0, tk.END)
        for c in self.curves:
            self.curve_listbox.insert(tk.END, f"{c['name']} ({os.path.basename(c['file'])})")
        if self.file_names:
            self.file_listbox.selection_set(0)
            self.on_file_select()
        else:
            self.df = None
            self.update_preview()
        self.plot_preview()

    def on_file_select(self, event=None):
        selected = self.file_listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        filename = self.file_names[idx]
        self.current_file = filename
        self.df = self.files[filename]
        self.update_preview()

    def update_preview(self, max_rows=20):
        for row in self.tree.get_children():
            self.tree.delete(row)
        if self.df is None:
            return
        cols = list(self.df.columns)
        self.tree["columns"] = cols
        self.tree["show"] = "headings"
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
        for i, row in self.df.head(max_rows).iterrows():
            values = [str(v) for v in row.values]
            self.tree.insert("", "end", values=values)
        if len(self.df) > max_rows:
            self.tree.insert("", "end", values=["... 共 {} 行，仅显示前 {} 行 ...".format(len(self.df), max_rows)])

    # ---------- 曲线编辑 ----------
    def edit_curve(self, event=None):
        selected = self.curve_listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        curve_data = self.curves[idx].copy()
        self.show_curve_dialog(edit_idx=idx, initial_data=curve_data)

    # ---------- 添加/编辑曲线对话框（公共方法） ----------
    def show_curve_dialog(self, edit_idx=None, initial_data=None, auto_assign=None, clear_table=None):
        is_edit = edit_idx is not None
        dialog = tk.Toplevel(self.master)
        dialog.title("编辑曲线" if is_edit else "添加曲线")
        dialog.geometry("700x650")
        dialog.transient(self.master)
        dialog.grab_set()

        # 文件选择
        ttk.Label(dialog, text="源文件:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        file_var = tk.StringVar()
        file_combo = ttk.Combobox(dialog, textvariable=file_var, values=self.file_names, state="readonly", width=40)
        file_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        if self.file_names:
            if is_edit and initial_data['file'] in self.file_names:
                file_combo.set(initial_data['file'])
            else:
                file_combo.set(self.file_names[0])

        # 图例名称
        ttk.Label(dialog, text="图例名称:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        name_var = tk.StringVar()
        if is_edit:
            name_var.set(initial_data['name'])
        ttk.Entry(dialog, textvariable=name_var, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Y列
        ttk.Label(dialog, text="Y列 (数值列):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        y_col_var = tk.StringVar()
        y_combo = ttk.Combobox(dialog, textvariable=y_col_var, state="readonly", width=40)
        y_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        if is_edit:
            y_col_var.set(initial_data['y_col'])

        # ---------- 行标识选择区域 ----------
        select_frame = ttk.LabelFrame(dialog, text="选择行标识（按住 Ctrl 多选）", padding=5)
        select_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        row_id_listbox = tk.Listbox(select_frame, height=8, width=30, selectmode="multiple")
        row_id_listbox.pack(side="left", padx=5, pady=5)

        btn_sel_frame = ttk.Frame(select_frame)
        btn_sel_frame.pack(side="left", padx=5, pady=5)
        ttk.Button(btn_sel_frame, text="添加选中点 →", command=lambda: add_selected()).pack(pady=2)
        ttk.Button(btn_sel_frame, text="自动分配", command=auto_assign).pack(pady=2)
        ttk.Button(btn_sel_frame, text="清空表格", command=clear_table).pack(pady=2)

        # ---------- 映射表格 ----------
        table_frame = ttk.LabelFrame(dialog, text="映射表格（行标识 → x 坐标，按 Enter 跳转下一行）", padding=5)
        table_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        dialog.grid_rowconfigure(4, weight=1)

        canvas = tk.Canvas(table_frame, height=200)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        row_widgets = []  # (label, entry, del_btn)

        # ---------- 内部函数 ----------
        def update_y_cols(*args):
            fname = file_var.get()
            if fname in self.files:
                df = self.files[fname]
                numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
                y_combo['values'] = numeric_cols
                if is_edit and initial_data['y_col'] in numeric_cols:
                    y_combo.set(initial_data['y_col'])
                elif numeric_cols:
                    y_combo.set(numeric_cols[0])
                else:
                    y_combo.set('')
                first_col = df.columns[0] if len(df.columns) > 0 else ''
                row_id_listbox.delete(0, tk.END)
                if first_col:
                    for item in df[first_col].astype(str).tolist():
                        row_id_listbox.insert(tk.END, item)
                # 如果是编辑模式，填充已有映射
                if is_edit and fname == initial_data['file']:
                    clear_table()
                    for row_id, x_val in initial_data['mapping'].items():
                        add_table_row(row_id)
                        row_widgets[-1][1].delete(0, tk.END)
                        row_widgets[-1][1].insert(0, str(x_val))

        def add_selected():
            selected = row_id_listbox.curselection()
            if not selected:
                messagebox.showwarning("警告", "请先选择行标识！")
                return
            selected_ids = [row_id_listbox.get(i) for i in selected]
            existing = {w[0].cget("text") for w in row_widgets}
            new_ids = [sid for sid in selected_ids if sid not in existing]
            if not new_ids:
                messagebox.showinfo("提示", "所选行标识已全部在表格中。")
                return
            for sid in new_ids:
                add_table_row(sid)

        def add_table_row(row_id):
            row_frame = ttk.Frame(scrollable_frame)
            row_frame.pack(fill="x", padx=2, pady=1)

            lbl = ttk.Label(row_frame, text=row_id, width=25)
            lbl.pack(side="left", padx=2)

            entry = ttk.Entry(row_frame, width=15)
            entry.pack(side="left", padx=2)
            entry.bind("<Return>", lambda e, rf=row_frame: on_enter(rf))
            entry.bind("<FocusIn>", lambda e: e.widget.select_range(0, tk.END))

            del_btn = ttk.Button(row_frame, text="×", width=2,
                                 command=lambda rf=row_frame, r_id=row_id: delete_row(rf, r_id))
            del_btn.pack(side="left", padx=2)

            row_widgets.append((lbl, entry, del_btn))
            entry.focus_set()

        def delete_row(row_frame, row_id):
            for idx, (lbl, entry, btn) in enumerate(row_widgets):
                if lbl.cget("text") == row_id:
                    row_frame.destroy()
                    del row_widgets[idx]
                    break

        def clear_table():
            for lbl, entry, btn in row_widgets:
                lbl.master.destroy()
            row_widgets.clear()

        def on_enter(current_frame):
            for i, (lbl, entry, btn) in enumerate(row_widgets):
                if entry.master == current_frame:
                    if i + 1 < len(row_widgets):
                        next_entry = row_widgets[i + 1][1]
                        next_entry.focus_set()
                        next_entry.select_range(0, tk.END)
                    break

        def auto_assign():
            clear_table()
            ids = row_id_listbox.get(0, tk.END)
            if not ids:
                messagebox.showwarning("警告", "行标识列表为空！")
                return
            for i, sid in enumerate(ids, start=1):
                add_table_row(sid)
                row_widgets[-1][1].delete(0, tk.END)
                row_widgets[-1][1].insert(0, str(float(i)))
            messagebox.showinfo("完成", f"已自动分配 {len(ids)} 个点的 x 值。")

        # 颜色选择
        color_var = tk.StringVar()
        if is_edit:
            color_var.set(initial_data['color'])
        else:
            color_var.set(self.get_auto_color())
        ttk.Label(dialog, text="颜色:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
        color_frame = ttk.Frame(dialog)
        color_frame.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        color_label = ttk.Label(color_frame, text="●", foreground=color_var.get())
        color_label.pack(side="left")

        def pick_color():
            color = colorchooser.askcolor(title="选择颜色")[1]
            if color:
                color_var.set(color)
                color_label.config(foreground=color)

        ttk.Button(color_frame, text="选择颜色", command=pick_color).pack(side="left", padx=5)

        # 确认按钮
        def confirm():
            fname = file_var.get()
            if not fname:
                messagebox.showwarning("警告", "请选择源文件！")
                return
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "请输入图例名称！")
                return
            y_col = y_col_var.get()
            if not y_col:
                messagebox.showwarning("警告", "请选择Y列！")
                return
            if not row_widgets:
                messagebox.showwarning("警告", "映射表格为空！")
                return
            mapping = {}
            for lbl, entry, btn in row_widgets:
                row_id = lbl.cget("text")
                x_str = entry.get().strip()
                if not x_str:
                    messagebox.showwarning("警告", f"行 '{row_id}' 的 x 值未填写！")
                    return
                try:
                    x_val = float(x_str)
                except ValueError:
                    messagebox.showerror("错误", f"行 '{row_id}' 的 x 值不是有效数字！")
                    return
                mapping[row_id] = x_val

            df = self.files[fname]
            first_col = df.columns[0] if len(df.columns) > 0 else ''
            if first_col:
                valid_ids = set(df[first_col].astype(str))
                invalid = [k for k in mapping.keys() if k not in valid_ids]
                if invalid:
                    if not messagebox.askyesno("警告", f"以下行标识在数据中不存在：{invalid}\n是否继续？"):
                        return
            else:
                messagebox.showerror("错误", "数据文件没有列！")
                return

            color = color_var.get()

            new_curve = {
                'name': name,
                'file': fname,
                'y_col': y_col,
                'mapping': mapping.copy(),
                'color': color
            }

            if is_edit:
                self.curves[edit_idx] = new_curve
                self.curve_listbox.delete(edit_idx)
                self.curve_listbox.insert(edit_idx, f"{name} ({os.path.basename(fname)})")
            else:
                self.curves.append(new_curve)
                self.curve_listbox.insert(tk.END, f"{name} ({os.path.basename(fname)})")
                self.auto_color_idx += 1

            dialog.destroy()
            self.plot_preview()

        ttk.Button(dialog, text="确定", command=confirm).grid(row=6, column=0, padx=5, pady=10)
        ttk.Button(dialog, text="取消", command=dialog.destroy).grid(row=6, column=1, padx=5, pady=10)

        file_var.trace('w', update_y_cols)
        update_y_cols()

    def add_curve(self):
        if not self.files:
            messagebox.showwarning("警告", "请先添加至少一个数据文件！")
            return
        self.show_curve_dialog(edit_idx=None, initial_data=None)

    def delete_curve(self):
        selected = self.curve_listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        self.curve_listbox.delete(idx)
        del self.curves[idx]
        self.plot_preview()

    def clear_curves(self):
        self.curves.clear()
        self.curve_listbox.delete(0, tk.END)
        self.auto_color_idx = 0
        self.plot_preview()

    def get_auto_color(self):
        color = self.colors[self.auto_color_idx % len(self.colors)]
        self.auto_color_idx += 1
        return color

    # ---------- 绘图核心（供预览和保存共用） ----------
    def draw_curves_on_ax(self, ax, m):
        """在给定的 Axes 上绘制所有曲线，返回是否有有效数据"""
        if not self.curves:
            return False

        for curve in self.curves:
            fname = curve['file']
            if fname not in self.files:
                continue
            df = self.files[fname]
            first_col = df.columns[0] if len(df.columns) > 0 else ''
            if not first_col:
                continue
            y_col = curve['y_col']
            mapping = curve['mapping']
            color = curve['color']

            xs = []
            ys = []
            for key, x_val in mapping.items():
                row = df[df[first_col].astype(str) == key]
                if row.empty:
                    continue
                try:
                    y_val = pd.to_numeric(row.iloc[0][y_col], errors='raise')
                except (ValueError, TypeError):
                    continue
                if pd.isna(y_val):
                    continue
                xs.append(x_val)
                ys.append(y_val)

            if len(xs) < 2:
                continue

            sorted_idx = np.argsort(xs)
            xs_sorted = np.array(xs)[sorted_idx]
            ys_sorted = np.array(ys)[sorted_idx]

            ax.scatter(xs_sorted, ys_sorted, marker='_', s=1200, linewidths=3, color=color, label=curve['name'])
            x2 = []
            y2 = []
            for x, y in zip(xs_sorted, ys_sorted):
                x2.extend([x - m, x + m])
                y2.extend([y, y])
            ax.plot(x2, y2, '--', color=color, linewidth=1.5, alpha=0.7)

        ax.set_xlabel("X")
        ax.set_ylabel("Intensity")
        ax.grid(True, linestyle=':', alpha=0.6)
        if len(self.curves) > 0:
            ax.legend(loc='best', frameon=True, fancybox=True, framealpha=0.9, edgecolor='gray')
        return True

    # ---------- 预览 ----------
    def plot_preview(self):
        self.ax.clear()
        if not self.curves:
            self.ax.set_title("请添加曲线")
            self.canvas.draw()
            return
        m = self.m_var.get()
        has_data = self.draw_curves_on_ax(self.ax, m)
        if not has_data:
            self.ax.set_title("没有有效数据可绘制")
        self.canvas.draw()

    # ---------- 保存（弹出独立窗口） ----------
    def save_figure(self):
        if not self.curves:
            messagebox.showwarning("警告", "没有曲线可保存！")
            return

        m = self.m_var.get()
        fig_new, ax_new = plt.subplots(figsize=(8, 6))
        has_data = self.draw_curves_on_ax(ax_new, m)
        if not has_data:
            plt.close(fig_new)
            messagebox.showwarning("警告", "没有有效数据可绘制！")
            return

        fig_new.tight_layout()
        messagebox.showinfo("提示",
                            "图形窗口已弹出，您可以使用工具栏调整图形（缩放、平移、调整子图等），\n"
                            "然后使用窗口的保存按钮（或 Ctrl+S）保存为所需格式。")
        plt.show()


if __name__ == "__main__":
    root = tk.Tk()
    app = EnergyProfileApp(root)
    root.mainloop()
