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
    
    def create_ansi_version_file(self, folder_path, target_version):
        """
        创建ANSI编码的version文件
        1. 删除原本的version文件（如果存在）
        2. 创建新的version文件
        3. 写入版本号和一个不可见的ANSI字符（0xAD）
        4. 用GBK编码保存（Windows中文系统的ANSI）
        """
        print(f"\n=== 创建ANSI编码的version文件 ===")
        version_path = os.path.join(folder_path, "version")
        
        try:
            # 步骤1: 删除原本的version文件（如果存在）
            if os.path.exists(version_path):
                try:
                    os.remove(version_path)
                    print(f"  已删除旧的version文件")
                except Exception as e:
                    print(f"  删除旧version文件失败: {e}")
                    return False
            
            # 步骤2: 创建新的version文件并写入内容
            # 准备要写入的内容：版本号 + 不可见ANSI字符
            # 使用0xAD（软连字符）作为不可见字符
            invisible_char = chr(0xAD)
            content = target_version + invisible_char
            
            # 用GBK编码写入文件
            with open(version_path, 'wb') as f:
                gbk_bytes = content.encode('gbk')
                f.write(gbk_bytes)
            
            print(f"  已创建新的version文件")
            print(f"  写入内容: '{target_version}' + 不可见字符(0xAD)")
            print(f"  文件大小: {len(gbk_bytes)} 字节")
            print(f"  文件字节(HEX): {gbk_bytes.hex(' ')}")
            
            # 步骤3: 验证文件是否成功创建
            if os.path.exists(version_path):
                print(f"  ✅ version文件创建成功")
                
                # 验证文件内容
                return self.verify_version_file(version_path, target_version)
            else:
                print(f"  ❌ version文件创建失败，文件不存在")
                return False
            
        except Exception as e:
            print(f"  创建version文件失败: {e}")
            return False
    
    def verify_version_file(self, file_path, expected_version):
        """
        验证version文件内容和编码
        """
        print(f"\n=== 验证version文件 ===")
        
        try:
            # 读取文件原始字节
            with open(file_path, 'rb') as f:
                raw_bytes = f.read()
            
            print(f"  文件大小: {len(raw_bytes)} 字节")
            print(f"  字节内容(HEX): {raw_bytes.hex(' ')}")
            
            # 检查是否有BOM
            if raw_bytes.startswith(b'\xef\xbb\xbf'):
                print(f"  ❌ 检测到UTF-8 BOM")
                return False
            
            # 尝试用GBK解码
            try:
                content = raw_bytes.decode('gbk')
                print(f"  ✅ 可以用GBK解码")
                
                # 移除不可见字符并检查版本号
                visible_content = content.replace(chr(0xAD), '')
                print(f"  可见内容（移除不可见字符后）: '{visible_content}'")
                
                if visible_content == expected_version:
                    print(f"  ✅ 版本号正确: {expected_version}")
                    
                    # 检查记事本识别情况
                    notepad_result = self.test_notepad_detection(file_path)
                    print(f"  记事本识别为: {notepad_result}")
                    
                    if notepad_result == "ansi":
                        print(f"  ✅ 记事本将识别为ANSI编码")
                    else:
                        print(f"  ⚠️ 注意：记事本可能识别为{notepad_result}")
                    
                    return True
                else:
                    print(f"  ❌ 版本号不正确，期望: '{expected_version}'，实际: '{visible_content}'")
                    return False
                    
            except Exception as e:
                print(f"  ❌ 无法用GBK解码: {e}")
                return False
                
        except Exception as e:
            print(f"  验证失败: {e}")
            return False
    
    def test_notepad_detection(self, file_path):
        """
        测试记事本如何检测文件编码
        """
        with open(file_path, 'rb') as f:
            raw_bytes = f.read()
        
        if not raw_bytes:
            return "empty"
        
        # 检查BOM
        if raw_bytes.startswith(b'\xef\xbb\xbf'):
            return "utf-8"
        
        # 检查是否全是ASCII
        if all(b < 128 for b in raw_bytes):
            return "ansi"
        
        # 尝试UTF-8解码
        try:
            raw_bytes.decode('utf-8')
            # 进一步检查：如果包含UTF-8多字节序列
            has_multibyte = any(b >= 0xC0 for b in raw_bytes)
            if has_multibyte:
                return "utf-8"
            else:
                return "ansi"
        except UnicodeDecodeError:
            return "ansi"
    
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
                            if self.convert_file_to_gbk(full_path):
                                processed_files.append(base_name)
                            file_found = True
                            break
                if not file_found:
                    print(f"未找到文件: {base_name}")
            
            # 3. 创建ANSI编码的version文件（先删除旧文件，再创建新文件）
            print(f"\n=== 处理version文件 ===")
            version_path = os.path.join(self.selected_folder, "version")
            
            # 检查旧的version文件是否存在
            if os.path.exists(version_path):
                print(f"  检测到已存在的version文件，将先删除后重新创建")
            
            version_created = self.create_ansi_version_file(self.selected_folder, target_version)
            
            if version_created:
                processed_files.append("version")
                print(f"\n✅ version文件处理成功")
                
                # 最终验证
                try:
                    with open(version_path, 'r', encoding='gbk') as f:
                        content = f.read()
                    
                    # 移除不可见字符
                    visible_content = content.replace(chr(0xAD), '')
                    print(f"最终version文件内容: '{visible_content}'")
                    print(f"文件编码: ANSI (GBK)")
                    print(f"记事本将显示为: ANSI编码")
                    
                except Exception as e:
                    print(f"读取version文件失败: {e}")
            else:
                print(f"\n❌ version文件创建失败")
            
            print(f"\n=== 处理结果汇总 ===")
            print(f"已处理文件: {', '.join(processed_files) if processed_files else '无'}")
            
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
