#!/usr/bin/env python3
"""
Excel合并工具 - 多个Excel文件合并为一个
"""
import sys, os, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

class App:
    def __init__(self, root):
        self.root = root
        root.title("Excel合并工具 v1.0")
        root.geometry("600x450")
        self.files = []
        self.build_ui()
    
    def build_ui(self):
        f = tk.Frame(self.root, bg="#388e3c", height=50)
        f.pack(fill="x")
        tk.Label(f, text="📊 Excel合并工具", font=("Arial",14,"bold"),
                 fg="white", bg="#388e3c").pack(pady=12)
        
        main = tk.Frame(self.root, padx=15, pady=10)
        main.pack(fill="both", expand=True)
        
        bf = tk.Frame(main)
        bf.pack(fill="x", pady=5)
        tk.Button(bf, text="添加Excel文件", command=self.add_files,
                  bg="#388e3c", fg="white", padx=12).pack(side="left", padx=5)
        tk.Button(bf, text="清空列表", command=self.clear,
                  bg="#d9534f", fg="white", padx=12).pack(side="left", padx=5)
        
        self.lb = tk.Listbox(main, font=("Consolas",10), bg="#e8f5e9", height=10)
        self.lb.pack(fill="both", expand=True, pady=10)
        
        tk.Button(main, text="🚀 合并并保存", command=self.merge,
                  bg="#4caf50", fg="white", font=("Arial",11,"bold"),
                  padx=20, pady=8).pack(pady=10)
        
        self.status = tk.Label(main, text="请添加要合并的Excel文件",
                               font=("Arial",10), fg="gray")
        self.status.pack()
    
    def add_files(self):
        fs = filedialog.askopenfilenames(title="选择Excel文件",
             filetypes=[("Excel","*.xlsx *.xls")])
        for f in fs:
            if f not in self.files:
                self.files.append(f)
                self.lb.insert("end", Path(f).name)
        self.status.config(text=f"已添加 {len(self.files)} 个文件")
    
    def clear(self):
        self.files.clear()
        self.lb.delete(0, "end")
        self.status.config(text="列表已清空")
    
    def merge(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加Excel文件")
            return
        if not HAS_PANDAS:
            messagebox.showerror("缺少依赖", "请运行：pip install pandas openpyxl")
            return
        
        out = filedialog.asksaveasfilename(title="保存合并后的Excel",
             defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if not out: return
        
        try:
            dfs = [pd.read_excel(f) for f in self.files]
            merged = pd.concat(dfs, ignore_index=True)
            merged.to_excel(out, index=False)
            messagebox.showinfo("完成", f"成功合并 {len(self.files)} 个文件！\n保存至：{out}")
            self.status.config(text=f"✅ 完成！共 {len(merged)} 行")
        except Exception as e:
            messagebox.showerror("错误", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
