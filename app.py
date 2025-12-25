import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from docx import Document
import re
import os
import threading

def process_single_file(file_path):
    """处理单个文件的逻辑，支持处理软回车"""
    try:
        doc = Document(file_path)
        new_doc = Document()
        has_chinese_content = False
        
        # 遍历所有段落
        for para in doc.paragraphs:
            # 1. 获取段落文本
            full_text = para.text
            
            # 2. 关键步骤：按换行符(\n)将段落切分为“子行”
            # Word中的软回车(Shift+Enter)在python-docx里通常表现为 \n
            sub_lines = full_text.split('\n')
            
            for line in sub_lines:
                clean_line = line.strip()
                # 3. 如果这一行包含中文，才保留
                if re.search(r'[\u4e00-\u9fa5]', clean_line):
                    new_doc.add_paragraph(clean_line)
                    has_chinese_content = True
        
        # --- 如果文档里有表格，也处理一下表格 (可选) ---
        for table in doc.tables:
            for row in table.rows:
                for cell in table.cells:
                    # 同样的逻辑处理表格里的字
                    if re.search(r'[\u4e00-\u9fa5]', cell.text):
                        pass # 这里简单处理，表格结构复杂，很难完美拆分，暂不删除表格内容
                    else:
                        cell.text = "" # 如果单元格全是英文，清空它

        # 构造新文件名
        dir_name = os.path.dirname(file_path)
        base_name = os.path.basename(file_path)
        name_part = os.path.splitext(base_name)[0]
        new_path = os.path.join(dir_name, f"{name_part}_纯中文版.docx")
        
        new_doc.save(new_path)
        
        if not has_chinese_content:
            return True, f"完成 (警告:原文档无中文): {base_name}"
        return True, f"成功: {base_name}"
        
    except Exception as e:
        return False, f"失败 {os.path.basename(file_path)}: {str(e)}"

# 下面的界面代码保持不变，直接复制之前的即可，或者只替换上面这个函数
# 为了方便你复制，这里把界面部分也补全：

def run_batch_process(files):
    total = len(files)
    success_count = 0
    progress_bar['maximum'] = total
    progress_bar['value'] = 0
    log_text.config(state='normal')
    log_text.delete(1.0, tk.END)
    
    for i, file_path in enumerate(files):
        status_label.config(text=f"处理中: {os.path.basename(file_path)}")
        success, msg = process_single_file(file_path)
        if success:
            success_count += 1
            log_text.insert(tk.END, f"[√] {msg}\n", 'success')
        else:
            log_text.insert(tk.END, f"[X] {msg}\n", 'error')
        log_text.see(tk.END)
        progress_bar['value'] = i + 1
        root.update_idletasks()
        
    log_text.config(state='disabled')
    status_label.config(text="全部完成")
    btn.config(state='normal')
    messagebox.showinfo("完成", f"处理结束！\n成功：{success_count}")

def select_files():
    file_paths = filedialog.askopenfilenames(title="选择文件", filetypes=[("Word", "*.docx")])
    if file_paths:
        btn.config(state='disabled')
        threading.Thread(target=run_batch_process, args=(file_paths,)).start()

root = tk.Tk()
root.title("Word强力去英文工具")
root.geometry("500x450")
tk.Label(root, text="Word文档中英过滤 (强力版)", font=("微软雅黑", 16, "bold")).pack(pady=10)
btn = tk.Button(root, text="📂 批量选择文件", font=("微软雅黑", 11), command=select_files, bg="#0078D7", fg="white", height=2)
btn.pack(pady=10)
progress_bar = ttk.Progressbar(root, length=400)
progress_bar.pack(pady=10)
status_label = tk.Label(root, text="准备就绪")
status_label.pack()
log_text = tk.Text(root, height=10)
log_text.pack(padx=20, pady=10)
log_text.tag_config('success', foreground='green')
log_text.tag_config('error', foreground='red')
root.mainloop()