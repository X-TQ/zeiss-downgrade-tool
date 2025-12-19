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
            except:
                pass

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
    
    def convert_file_to_gbk(self, file_path):
        """转换普通文件为GBK编码"""
        try:
            # 先尝试用各种可能的编码读取
            encodings_to_try = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'latin-1', 'cp1252']
            content = None
            used_encoding = None
            
            for enc in encodings_to_try:
                try:
                    with open(file_path, 'r', encoding=enc) as f:
                        content = f.read()
                    used_encoding = enc
                    print(f"  成功使用 {enc} 编码读取 {os.path.basename(file_path)}")
                    break
                except Exception as e:
                    continue
            
            # 如果以上方法都失败，尝试二进制读取
            if content is None:
                print(f"  无法用标准编码读取 {os.path.basename(file_path)}，尝试二进制解码")
                with open(file_path, 'rb') as f:
                    raw_data = f.read()
                
                # 尝试UTF-8解码
                try:
                    content = raw_data.decode('utf-8')
                    used_encoding = 'utf-8 (binary)'
                except:
                    try:
                        content = raw_data.decode('gbk')
                        used_encoding = 'gbk (binary)'
                    except:
                        # 最后尝试latin-1，它总是能解码
                        content = raw_data.decode('latin-1', errors='ignore')
                        used_encoding = 'latin-1 (binary)'
            
            # 用GBK编码写入文件
            with open(file_path, 'w', encoding='gbk', errors='ignore') as f:
                f.write(content)
            
            print(f"  已将 {os.path.basename(file_path)} 转换为GBK编码 (原编码: {used_encoding})")
            return True
            
        except Exception as e:
            print(f"  转换 {os.path.basename(file_path)} 失败: {e}")
            return False
    
    def convert_version_to_ansi(self, folder_path, target_version):
        """
        确保version文件为ANSI编码（GBK）
        特别注意：如果原始文件是UTF-8，需要确保转换为ANSI（GBK）
        """
        print(f"\n=== 转换version文件为ANSI编码 ===")
        version_path = os.path.join(folder_path, "version")
        
        try:
            # 1. 检查是否存在现有的version文件
            existing_version = None
            if os.path.exists(version_path):
                # 尝试读取现有文件内容
                try:
                    with open(version_path, 'r', encoding='utf-8') as f:
                        existing_version = f.read().strip()
                    print(f"  检测到UTF-8编码的version文件，内容: '{existing_version}'")
                except:
                    try:
                        with open(version_path, 'r', encoding='gbk') as f:
                            existing_version = f.read().strip()
                        print(f"  检测到GBK编码的version文件，内容: '{existing_version}'")
                    except:
                        print(f"  无法读取现有version文件，将创建新文件")
            
            # 2. 确保使用GBK编码写入（Windows下的ANSI编码）
            # 注意：这里不添加BOM，因为ANSI编码没有BOM
            with open(version_path, 'w', encoding='gbk', errors='strict') as f:
                f.write(target_version)
            
            print(f"  已创建ANSI编码的version文件")
            print(f"  写入的版本号: {target_version}")
            
            # 3. 验证文件编码
            verification = self.verify_file_encoding(version_path, "version")
            if verification == "gbk":
                print(f"  ✅ 验证成功：version文件已转换为ANSI/GBK编码")
                return True
            else:
                print(f"  ⚠️ 警告：version文件可能不是ANSI编码，检测为: {verification}")
                # 尝试修复：用二进制方式重新写入
                try:
                    with open(version_path, 'wb') as f:
                        # 使用gbk编码，不添加BOM
                        f.write(target_version.encode('gbk'))
                    print(f"  已用二进制方式重新写入GBK编码")
                    return True
                except:
                    return False
                
        except Exception as e:
            print(f"  创建ANSI编码version文件失败: {e}")
            return False
    
    def verify_file_encoding(self, file_path, filename):
        """
        验证文件的编码格式，返回检测到的编码
        """
        try:
            with open(file_path, 'rb') as f:
                raw_bytes = f.read(1024)  # 读取前1024字节
            
            if not raw_bytes:
                return "empty"
            
            # 检查BOM
            if raw_bytes.startswith(b'\xef\xbb\xbf'):
                return "utf-8-bom"
            elif raw_bytes.startswith(b'\xff\xfe'):
                return "utf-16-le"
            elif raw_bytes.startswith(b'\xfe\xff'):
                return "utf-16-be"
            
            # 检查是否为纯ASCII（兼容ANSI和UTF-8）
            is_ascii = all(b < 128 for b in raw_bytes)
            
            # 尝试用UTF-8解码
            try:
                raw_bytes.decode('utf-8')
                if is_ascii:
                    return "ascii (兼容utf-8和ansi)"
                else:
                    return "utf-8"
            except:
                # 尝试用GBK解码
                try:
                    raw_bytes.decode('gbk')
                    return "gbk"
                except:
                    return "unknown"
                    
        except Exception as e:
            return f"error: {str(e)}"
    
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
            
            # 2. 处理其他文件
            print(f"\n=== 处理 inspset, inspection, username 文件 ===")
            other_files = ["inspset", "inspection", "username"]
            processed_files = []
            
            for base_name in other_files:
                file_found = False
                for filename in os.listdir(self.selected_folder):
                    full_path = os.path.join(self.selected_folder, filename)
                    if os.path.isfile(full_path):
                        name_without_ext = os.path.splitext(filename)[0]
                        if name_without_ext.lower() == base_name.lower():
                            print(f"\n处理文件: {filename}")
                            # 先验证原始编码
                            original_encoding = self.verify_file_encoding(full_path, filename)
                            print(f"  原始编码: {original_encoding}")
                            
                            if self.convert_file_to_gbk(full_path):
                                # 验证转换后的编码
                                new_encoding = self.verify_file_encoding(full_path, filename)
                                print(f"  转换后编码: {new_encoding}")
                                processed_files.append(f"{base_name}")
                            file_found = True
                            break
                if not file_found:
                    print(f"未找到文件: {base_name}")
            
            # 3. 创建或转换version文件为ANSI编码
            print(f"\n=== 处理version文件 ===")
            version_path = os.path.join(self.selected_folder, "version")
            
            # 如果version文件已存在，先备份
            if os.path.exists(version_path):
                try:
                    backup_path = version_path + ".backup"
                    shutil.copy2(version_path, backup_path)
                    print(f"  已备份原始version文件到: {backup_path}")
                except:
                    print(f"  警告：无法备份version文件")
            
            version_success = self.convert_version_to_ansi(self.selected_folder, target_version)
            
            if version_success:
                processed_files.append("version")
                
                # 最终验证
                print(f"\n=== 最终验证 ===")
                if os.path.exists(version_path):
                    final_encoding = self.verify_file_encoding(version_path, "version")
                    print(f"version文件最终编码: {final_encoding}")
                    
                    # 尝试读取并显示内容
                    try:
                        with open(version_path, 'r', encoding='gbk') as f:
                            content = f.read().strip()
                        print(f"用GBK读取成功: '{content}'")
                        
                        # 确保内容正确
                        if content == target_version:
                            print(f"✅ version文件内容验证正确")
                        else:
                            print(f"⚠️ 警告：version文件内容不匹配，期望: '{target_version}'，实际: '{content}'")
                    except Exception as e:
                        print(f"无法用GBK读取version文件: {e}")
            else:
                print(f"\n❌ 错误：创建ANSI编码的version文件失败！")
            
            print(f"\n=== 处理结果汇总 ===")
            print(f"已处理文件: {', '.join(processed_files) if processed_files else '无'}")
            
            self.status_label.config(text="恭喜您，降级成功！！！", fg="green")
            messagebox.showinfo("完成", f"恭喜您，降级成功！！！\n已降级到版本: {target_version}\n\n注意：version文件已转换为ANSI编码")
            
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
