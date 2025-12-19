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
        self.set_window_icon()
        self.selected_folder = ""
        self.versions = self.generate_versions()
        self.create_widgets()
    
    def set_window_icon(self):
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
    
    def convert_file_to_gbk(self, file_path, target_content=None):
        """
        将文件内容读取并转换为GBK编码。
        如果提供了 target_content，则使用该内容覆盖原文件，然后转换编码。
        否则，读取原文件内容后转换编码。
        """
        try:
            # 如果要写入特定内容，则直接使用
            if target_content is not None:
                content_to_save = str(target_content)
                print(f"  将写入指定内容: '{content_to_save}'")
            else:
                # 否则，尝试读取原文件内容
                encodings_to_try = ['utf-8', 'gbk', 'gb2312', 'latin-1', 'cp1252', 'utf-8-sig']
                content_to_save = None
                for enc in encodings_to_try:
                    try:
                        with open(file_path, 'r', encoding=enc) as f:
                            content_to_save = f.read()
                        print(f"  成功读取原文件，编码: {enc}")
                        break
                    except:
                        continue
                if content_to_save is None:
                    with open(file_path, 'rb') as f:
                        raw_data = f.read()
                    try:
                        content_to_save = raw_data.decode('utf-8')
                    except:
                        try:
                            content_to_save = raw_data.decode('gbk')
                        except:
                            content_to_save = raw_data.decode('latin-1', errors='ignore')
                    print(f"  以二进制方式读取并解码文件。")
            
            # 最终使用 GBK 编码写入文件
            with open(file_path, 'w', encoding='gbk', errors='ignore') as f:
                f.write(content_to_save)
            print(f"  成功写入GBK编码。")
            return True
        except Exception as e:
            print(f"  转换失败 {os.path.basename(file_path)}: {e}")
            return False
    
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
                    print(f"已删除文件夹: geoactuals")
                except Exception as e:
                    print(f"删除 geoactuals 失败: {e}")
            
            # 2. 【核心修改】统一处理四个文件
            print(f"\n=== 统一处理 inspset, inspection, username, version 文件 ===")
            target_files = ["inspset", "inspection", "username", "version"]
            processed_files = []
            
            for base_name in target_files:
                file_found = False
                for filename in os.listdir(self.selected_folder):
                    full_path = os.path.join(self.selected_folder, filename)
                    if os.path.isfile(full_path):
                        name_without_ext = os.path.splitext(filename)[0]
                        if name_without_ext.lower() == base_name.lower():
                            print(f"\n处理文件: {filename} (匹配: {base_name})")
                            file_found = True
                            
                            # *** 关键逻辑分支 ***
                            if base_name.lower() == "version":
                                # 对于 version 文件：先写入目标版本号，再进行GBK编码转换
                                print(f"  此为 version 文件，将内容设置为: {target_version}")
                                success = self.convert_file_to_gbk(full_path, target_content=target_version)
                            else:
                                # 对于其他三个文件：直接进行GBK编码转换（保留原内容）
                                success = self.convert_file_to_gbk(full_path)
                            
                            if success:
                                processed_files.append(base_name)
                            break # 找到一个匹配后即跳出内层循环
                
                # 如果没找到文件，并且是 version 文件，则创建它
                if not file_found and base_name.lower() == "version":
                    print(f"\n未找到 version 文件，将创建新文件。")
                    new_version_path = os.path.join(self.selected_folder, "version")
                    try:
                        # 直接创建并写入目标版本号，使用GBK编码
                        with open(new_version_path, 'w', encoding='gbk') as f:
                            f.write(target_version)
                        print(f"  已创建新文件 'version'，内容: '{target_version}' (GBK编码)")
                        processed_files.append("version")
                    except Exception as e:
                        print(f"  创建 version 文件失败: {e}")
                elif not file_found:
                    # 对于其他未找到的文件，仅记录
                    print(f"未找到文件: {base_name}")
            
            # 3. 最终验证与汇总
            print(f"\n=== 处理结果汇总 ===")
            print(f"已处理文件: {', '.join(processed_files) if processed_files else '无'}")
            
            # 特别验证 version 文件
            version_check_path = os.path.join(self.selected_folder, "version")
            if os.path.exists(version_check_path):
                try:
                    # 尝试用GBK读取，验证编码和内容
                    with open(version_check_path, 'r', encoding='gbk') as f:
                        final_content = f.read().strip()
                    print(f"Version 文件最终内容: '{final_content}' (可被GBK解码)")
                    # 尝试用UTF-8读取，如果失败则更确认是GBK
                    try:
                        with open(version_check_path, 'r', encoding='utf-8') as f:
                            utf8_content = f.read()
                        print("注意：文件也能用UTF-8解码，但GBK写入已确保。")
                    except UnicodeDecodeError:
                        print("确认：文件无法用UTF-8解码，编码应为ANSI/GBK。")
                except Exception as e:
                    print(f"验证 version 文件时出错: {e}")
            else:
                print("警告：最终未找到 version 文件。")
            
            self.status_label.config(text="恭喜您，降级成功！！！", fg="green")
            messagebox.showinfo("完成", f"恭喜您，降级成功！！！\n已降级到版本: {target_version}")
            
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
    root.geometry(f"{window_width}x{height}+{x}+{y}")
    root.mainloop()
