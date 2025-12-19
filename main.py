import os
import sys
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class ZeissDowngradeTool:
    def __init__(self, root):
        self.root = root
        self.root.title("蔡司降级工具 v1.0")
        self.root.geometry("500x380")
        self.root.resizable(False, False)
        
        # 设置窗口图标
        self.set_window_icon()
        
        self.selected_folder = ""
        self.versions = self.generate_versions()
        self.create_widgets()
    
    def set_window_icon(self):
        """设置窗口图标，兼容开发环境和打包后的EXE"""
        icon_name = "zeiss_icon.ico"
        icon_path = None
        
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
            icon_path = os.path.join(base_path, icon_name)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_path, icon_name)
        
        if icon_path and os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
                print(f"图标加载成功: {icon_path}")
            except Exception as e:
                print(f"警告: 无法加载图标文件 '{icon_path}'。原因: {e}")
        else:
            print(f"警告: 未找到图标文件 '{icon_name}'。请确保它存在于: {base_path}")
    
    def generate_versions(self):
        """生成降级版本列表（无奇数，每次递减0.2）"""
        versions = []
        current = 7.4
        while current >= 5.4:
            versions.append(f"{current:.1f}")
            current -= 0.2
            current = round(current, 1)
        return versions
    
    def create_widgets(self):
        title_label = tk.Label(self.root, text="蔡司降级工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        frame1 = tk.Frame(self.root)
        frame1.pack(pady=10, padx=20, fill='x')
        
        open_btn = tk.Button(frame1, text="选择降级文件夹", command=self.select_folder,
                           font=("Arial", 11), bg="#4CAF50", fg="white", height=2, width=20)
        open_btn.pack(pady=5)
        
        self.folder_label = tk.Label(frame1, text="未选择文件夹", font=("Arial", 9), fg="gray", wraplength=400)
        self.folder_label.pack(pady=5)
        
        frame2 = tk.Frame(self.root)
        frame2.pack(pady=10)
        
        version_label = tk.Label(frame2, text="选择降级版本：", font=("Arial", 10))
        version_label.pack(pady=5)
        
        self.version_var = tk.StringVar(value="7.4")
        version_combo = ttk.Combobox(frame2, textvariable=self.version_var,
                                   values=self.versions, state="readonly", font=("Arial", 10), width=10)
        version_combo.pack(pady=5)
        
        frame3 = tk.Frame(self.root)
        frame3.pack(pady=20, padx=20, fill='x')
        
        self.start_btn = tk.Button(frame3, text="开始降级", command=self.start_downgrade,
                                 font=("Arial", 12, "bold"), bg="#2196F3", fg="white",
                                 height=2, width=15, state="disabled")
        self.start_btn.pack()
        
        self.status_label = tk.Label(self.root, text="准备就绪", font=("Arial", 9), fg="green")
        self.status_label.pack(pady=10)
    
    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择要降级的文件夹")
        if folder_path:
            self.selected_folder = folder_path
            display_path = folder_path
            if len(folder_path) > 50:
                display_path = "..." + folder_path[-50:]
            self.folder_label.config(text=f"已选择: {display_path}", fg="blue")
            self.start_btn.config(state="normal")
            self.status_label.config(text="已选择文件夹，可以开始降级", fg="green")
    
    def convert_encoding_to_gbk(self, file_path):
        """将文件内容读取并转换为GBK编码（即Windows ANSI）"""
        try:
            encodings_to_try = ['utf-8', 'gbk', 'gb2312', 'latin-1', 'cp1252', 'utf-8-sig']
            content = None
            
            for enc in encodings_to_try:
                try:
                    with open(file_path, 'r', encoding=enc) as f:
                        content = f.read()
                    break
                except:
                    continue
            
            if content is None:
                with open(file_path, 'rb') as f:
                    raw_data = f.read()
                try:
                    content = raw_data.decode('utf-8')
                except:
                    try:
                        content = raw_data.decode('gbk')
                    except:
                        content = raw_data.decode('latin-1', errors='ignore')
            
            return content, True
        except Exception as e:
            print(f"转换文件编码失败 {file_path}: {e}")
            return None, False
    
    def start_downgrade(self):
        if not self.selected_folder:
            messagebox.showwarning("警告", "请先选择要降级的文件夹！")
            return
        
        target_version = self.version_var.get()
        
        confirm = messagebox.askyesno(
            "确认操作",
            f"确定要开始降级吗？\n\n文件夹: {os.path.basename(self.selected_folder)}\n降级到: {target_version}"
        )
        
        if not confirm:
            return
        
        try:
            self.status_label.config(text="正在处理...", fg="orange")
            self.root.update()
            
            # 1. 删除geoactuals文件夹
            geoactuals_path = os.path.join(self.selected_folder, "geoactuals")
            if os.path.exists(geoactuals_path):
                try:
                    shutil.rmtree(geoactuals_path)
                except:
                    pass
            
            # 2. 处理指定文件（inspset, inspection, username, version）
            target_files = ["inspset", "inspection", "username", "version"]
            files_processed = []
            
            for base_name in target_files:
                file_found = False
                for filename in os.listdir(self.selected_folder):
                    full_path = os.path.join(self.selected_folder, filename)
                    if os.path.isfile(full_path):
                        name_without_ext = os.path.splitext(filename)[0]
                        if name_without_ext.lower() == base_name.lower():
                            content, success = self.convert_encoding_to_gbk(full_path)
                            if success and content:
                                # 关键修改1：使用 'gbk' 代替 'ansi'
                                with open(full_path, 'w', encoding='gbk', errors='ignore') as f:
                                    f.write(content)
                                files_processed.append(base_name)
                                print(f"已转换编码: {filename} -> GBK")
                            file_found = True
                            break
                if not file_found:
                    print(f"未找到文件: {base_name}")
            
            # 3. 修改version文件内容（无论原文件是否存在或已被转换）
            # 关键修改2：确保version文件以GBK编码写入目标版本号
            version_updated = False
            version_file_path = os.path.join(self.selected_folder, "version")
            try:
                with open(version_file_path, 'w', encoding='gbk') as f:
                    f.write(target_version)
                version_updated = True
                print(f"已将version文件内容修改为: {target_version} (GBK编码)")
            except Exception as e:
                print(f"写入version文件失败: {e}")
            
            # 4. 验证并输出结果
            if version_updated:
                # 可选：再次读取验证编码
                try:
                    with open(version_file_path, 'r', encoding='gbk') as f:
                        verified_content = f.read().strip()
                    print(f"验证version文件内容: '{verified_content}'")
                except:
                    print("警告：无法以GBK编码重新读取version文件验证。")
            
            self.status_label.config(text="恭喜您，降级成功！！！", fg="green")
            messagebox.showinfo("完成", f"恭喜您，降级成功！！！\n已降级到版本: {target_version}\n处理文件: {', '.join(files_processed) if files_processed else '无'}")
            
            self.start_btn.config(state="disabled")
            self.folder_label.config(text="未选择文件夹", fg="gray")
            self.selected_folder = ""
            
        except Exception as e:
            self.status_label.config(text="降级失败", fg="red")
            messagebox.showerror("错误", f"降级过程中出现错误:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ZeissDowngradeTool(root)
    
    window_width = 500
    window_height = 380
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    root.mainloop()
