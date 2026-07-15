# 化学电离质谱（CIMS）数据处理工具包
# CIMS Data Analysis Toolkit for Iodide Chemical Ionization Mass Spectrometry

[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## 简介 | Introduction

**中文**：  
本工具包专为化学电离质谱（CIMS）实验数据设计，提供 **一站式的图形界面（GUI）数据处理解决方案**。涵盖从原始质谱数据到发表级图表的完整流程，包括：

- **峰强度提取**：自动搜索指定 m/z 附近的最高峰强度，生成汇总表。
- **数据归一化**：支持多分子列同时归一化（分子列分别除以分母列之和）。
- **模型拟合与评估**：内置 17 种数学模型，自动拟合两列数据并输出最佳模型及公式。
- **能垒图绘制**：交互式加载多个数据文件，自定义 x 坐标映射，绘制带台阶的散点图，支持多曲线叠加。

所有工具均以 **纯 GUI 形式** 运行，无需编写代码，适合科研人员快速处理实验数据。

---

**English**:  
This toolkit is specifically designed for **chemical ionization mass spectrometry (CIMS)** data, offering a **one-stop GUI-based data processing solution**. It covers the entire workflow from raw mass spectra to publication-ready figures, including:

- **Peak intensity extraction**: Automatically searches for the highest intensity near specified m/z values and generates a summary table.
- **Data normalization**: Supports multi-column normalization (each numerator column divided by the sum of denominator columns).
- **Model fitting & evaluation**: Built‑in 17 mathematical models for automatic fitting of two-column data, outputting the best model and equation.
- **Energy profile plotting**: Interactive loading of multiple data files with custom x‑coordinate mapping, producing step‑like scatter plots with multiple curve overlays.

All tools are **pure GUI‑driven** – no coding required, making them ideal for researchers needing quick data processing.

---

## 功能概览 | Feature Overview

| 工具 (Tool) | 功能 (Function) | 输入 (Input) | 输出 (Output) |
|------------|----------------|-------------|--------------|
| **1-质谱寻峰程序** (Peak Finder) | 在指定文件夹所有 `.xlsx` 文件中搜索目标 m/z 附近的最高峰强度，生成汇总表。 | 目标 m/z 列表（可修改代码）、数据文件夹 | `peak_intensity_summary.xlsx`（各文件对应峰强度） |
| **2-各种拟合程序** (Fitting Tool) | 对两列数据进行 17 种模型拟合，按调整 R² 排序，输出最佳模型公式和图形。 | 两列数据的 `.xlsx` 文件（第一列为 x，第二列为 y） | 拟合结果 `.txt` + 拟合曲线图 `.png` |
| **3-归一化工具** (Normalization) | 对汇总表进行多分子列归一化：选定分子列分别除以分母列之和，自动插入新列并保留原列顺序。 | `peak_intensity_summary.xlsx` 或任意表格 | 新 `.xlsx` 文件（含归一化列） |
| **4-能垒图绘制** (Energy Profile Plotter) | 加载多个归一化表格，通过交互式 GUI 自定义曲线的 x 坐标映射，绘制带“台阶”散点图，支持多曲线叠加、颜色选择、图例编辑。 | 任意 `.xlsx` 文件（多文件加载） | 内嵌预览 + 可保存的高清 PNG（弹出窗口自由调整） |

---

## 快速开始 | Quick Start

### 环境要求 | Requirements
- Python 3.7 或更高版本
- 依赖库：`pandas`, `numpy`, `matplotlib`, `scipy`, `openpyxl`, `tkinter`（通常内置）

### 安装依赖 | Install Dependencies
```bash
pip install pandas numpy matplotlib scipy openpyxl
```

### 致谢 | Acknowledgments
中文：
本工具包的开发得到了 DeepSeek 的深度参与和帮助。从代码框架设计、GUI 交互逻辑、错误调试，到功能扩展和性能优化，DeepSeek 提供了大量高效、准确的建议和解决方案，显著加速了开发进程。
在此向 DeepSeek 团队表示诚挚感谢！

English:
This toolkit was developed with substantial assistance from DeepSeek. From code architecture design, GUI interaction logic, bug fixing, to feature extensions and performance optimization, DeepSeek provided numerous efficient and accurate suggestions and solutions, greatly accelerating the development process.
Sincere thanks to the DeepSeek team!
