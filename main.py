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
    
    def convert_file_to_gbk(self, file_path):
        """将单个文件内容读取并转换为GBK编码（用于 inspset, inspection, username）"""
        try:
            encodings_to_try = ['utf-8', 'gbk', 'gb2312', 'latin-1', 'cp1252', 'utf-8-sig']
            content = None
            for enc in encodings_to_try:
                try:
                    with open(file_path, 'r', encoding=enc) as f:
                        content = f.read()
                    print(f"  读取成功，原编码: {enc}")
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
                print(f"  以二进制方式读取并解码。")
            
            # 使用 errors='ignore' 忽略无法编码的字符
            with open(file_path, 'w', encoding='gbk', errors='ignore') as f:
                f.write(content)
            print(f"  转换成功 -> GBK编码")
            return True
        except Exception as e:
            print(f"  转换失败: {e}")
            return False
    
    def create_version_file_gbk(self, folder_path, target_version):
        """
        【核心解决方案】创建GBK编码的version文件（二进制方式）
        完全绕过文本编码层，确保写入的就是GBK字节数据
        """
        print(f"\n=== 强制创建GBK编码的version文件 ===")
        
        # 1. 先删除任何已存在的version文件
        deleted = False
        for filename in os.listdir(folder_path):
            full_path = os.path.join(folder_path, filename)
            if os.path.isfile(full_path):
                name_without_ext = os.path.splitext(filename)[0]
                if name_without_ext.lower() == "version":
                    try:
                        os.remove(full_path)
                        print(f"  已删除旧文件: {filename}")
                        deleted = True
                    except Exception as e:
                        print(f"  删除失败 {filename}: {e}")
        
        # 2. 创建新的version文件（二进制模式）
        new_version_path = os.path.join(folder_path, "version")
        try:
            # 核心步骤：将字符串明确编码为GBK字节，然后用二进制模式写入
            gbk_bytes = target_version.encode('gbk', errors='ignore')
            print(f"  将版本号 '{target_version}' 编码为GBK字节: {gbk_bytes.hex()}")
            
            with open(new_version_path, 'wb') as f:  # 'wb' 表示二进制写入
                f.write(gbk_bytes)
            
            print(f"  已创建新文件 'version' (二进制GBK编码)")
            
            # 验证：尝试用GBK解码读取
            try:
                with open(new_version_path, 'rb') as f:
                    read_bytes = f.read()
                decoded_content = read_bytes.decode('gbk')
                print(f"  验证成功: 文件可被GBK解码，内容为 '{decoded_content}'")
            except Exception as e:
                print(f"  验证警告: {e}")
            
            return True
            
        except Exception as e:
            print(f"  创建文件失败: {e}")
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
            
            # 2. 处理 inspset, inspection, username 文件（原有逻辑）
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
                            if self.convert_file_to_gbk(full_path):
                                processed_files.append(base_name)
                            file_found = True
                            break
                if not file_found:
                    print(f"未找到文件: {base_name}")
            
            # 3. 【核心】使用专用函数创建version文件
            version_success = self.create_version_file_gbk(self.selected_folder, target_version)
            if version_success:
                processed_files.append("version")
            
            # 4. 最终验证
            print(f"\n=== 处理结果汇总 ===")
            print(f"已处理文件: {', '.join(processed_files) if processed_files else '无'}")
            
            # 特别验证version文件编码
            version_check_path = os.path.join(self.selected_folder, "version")
            if os.path.exists(version_check_path):
                # 方法1：尝试用不同编码读取
                test_results = []
                for encoding in ['gbk', 'utf-8', 'utf-16']:
                    try:
                        with open(version_check_path, 'r', encoding=encoding) as f:
                            content = f.read().strip()
                        test_results.append(f"{encoding}: 成功 ('{content}')")
                    except:
                        test_results.append(f"{encoding}: 失败")
                
                print(f"编码测试: {', '.join(test_results)}")
                
                # 方法2：检查BOM（字节顺序标记）
                with open(version_check_path, 'rb') as f:
                    first_bytes = f.read(3)
                bom_info = "无BOM"
                if first_bytes.startswith(b'\xef\xbb\xbf'):
                    bom_info = "有UTF-8 BOM"
                elif first_bytes.startswith(b'\xff\xfe') or first_bytes.startswith(b'\xfe\xff'):
                    bom_info = "有UTF-16 BOM"
                print(f"BOM检查: {bom_info}")
            
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
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.mainloop()
