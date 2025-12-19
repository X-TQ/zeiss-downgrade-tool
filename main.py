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
        创建ANSI编码的version文件，确保Windows记事本识别为ANSI
        通过在文件末尾添加特殊字节，让UTF-8解码失败，记事本就会使用ANSI
        """
        print(f"\n=== 创建ANSI编码的version文件 ===")
        version_path = os.path.join(folder_path, "version")
        
        try:
            # 将版本号转换为ASCII字节
            version_bytes = target_version.encode('ascii')
            
            # 添加一个无效的UTF-8序列
            # 0x81 是GBK编码中的一个字节，但不是有效的UTF-8起始字节
            # 单独的0x81无法用UTF-8解码，所以记事本会回退到ANSI
            invalid_utf8_bytes = b'\x81'
            
            with open(version_path, 'wb') as f:
                f.write(version_bytes)
                f.write(invalid_utf8_bytes)
            
            print(f"  已创建version文件")
            print(f"  原始字节: {version_bytes.hex(' ')} {invalid_utf8_bytes.hex(' ')}")
            print(f"  版本号: {target_version}")
            print(f"  添加了特殊字节以确保ANSI编码")
            
            return True
            
        except Exception as e:
            print(f"  创建version文件失败: {e}")
            return False
    
    def clean_version_file(self, folder_path, target_version):
        """
        清理version文件：删除乱码字符，只保留版本号
        但仍确保文件是ANSI编码
        """
        print(f"\n=== 清理version文件，删除乱码字符 ===")
        version_path = os.path.join(folder_path, "version")
        
        if not os.path.exists(version_path):
            print(f"  version文件不存在")
            return False
        
        try:
            # 读取文件原始字节
            with open(version_path, 'rb') as f:
                raw_bytes = f.read()
            
            print(f"  原始文件字节: {raw_bytes.hex(' ')}")
            print(f"  原始字节数: {len(raw_bytes)}")
            
            # 方法1：提取所有ASCII字符（版本号应该是ASCII）
            clean_bytes = bytearray()
            for b in raw_bytes:
                # ASCII可打印字符（包括数字和小数点）
                if 32 <= b <= 126:
                    clean_bytes.append(b)
            
            # 如果提取到了ASCII字符
            if clean_bytes:
                # 将字节转换回字符串
                clean_str = clean_bytes.decode('ascii')
                print(f"  提取的ASCII内容: '{clean_str}'")
                
                # 检查提取的内容是否以目标版本号开头
                if clean_str.startswith(target_version):
                    # 只保留版本号（去掉可能提取到的其他ASCII字符）
                    # 使用正则表达式提取版本号格式（如7.4, 7.2等）
                    import re
                    version_pattern = r'\d+\.\d+'
                    matches = re.findall(version_pattern, clean_str)
                    
                    if matches:
                        # 取第一个匹配的版本号
                        final_version = matches[0]
                        print(f"  匹配到的版本号: {final_version}")
                        
                        # 用GBK编码写入，确保是ANSI格式
                        with open(version_path, 'wb') as f:
                            # 使用GBK编码写入，但不添加任何特殊字节
                            gbk_bytes = final_version.encode('gbk')
                            f.write(gbk_bytes)
                        
                        print(f"  清理完成，最终版本号: {final_version}")
                        print(f"  最终字节: {gbk_bytes.hex(' ')}")
                        return True
                    else:
                        print(f"  警告：未找到版本号格式，直接写入目标版本号")
                else:
                    print(f"  警告：提取的内容不以目标版本号开头")
            
            # 方法2：如果方法1失败，直接写入目标版本号
            print(f"  使用方法2：直接写入目标版本号")
            with open(version_path, 'wb') as f:
                # 使用GBK编码写入
                gbk_bytes = target_version.encode('gbk')
                f.write(gbk_bytes)
            
            print(f"  直接写入目标版本号: {target_version}")
            print(f"  最终字节: {gbk_bytes.hex(' ')}")
            return True
            
        except Exception as e:
            print(f"  清理version文件失败: {e}")
            # 如果清理失败，尝试直接写入目标版本号
            try:
                with open(version_path, 'wb') as f:
                    f.write(target_version.encode('gbk'))
                print(f"  错误恢复：直接写入目标版本号")
                return True
            except:
                return False
    
    def test_notepad_detection(self, file_path):
        """
        测试记事本如何检测文件编码
        """
        print(f"\n=== 测试记事本编码检测 ===")
        
        with open(file_path, 'rb') as f:
            raw_bytes = f.read()
        
        if not raw_bytes:
            print("  文件为空")
            return "empty"
        
        # 记事本的简单检测逻辑：
        # 1. 检查BOM
        # 2. 尝试UTF-8解码，如果成功且不是纯ASCII，就认为是UTF-8
        # 3. 否则认为是ANSI
        
        # 检查BOM
        if raw_bytes.startswith(b'\xef\xbb\xbf'):
            print("  记事本会显示为: UTF-8 (因为有BOM)")
            return "utf-8"
        
        # 检查是否全是ASCII
        if all(b < 128 for b in raw_bytes):
            print("  记事本会显示为: ANSI (纯ASCII)")
            return "ansi"
        
        # 尝试UTF-8解码
        try:
            raw_bytes.decode('utf-8')
            # 进一步检查：如果包含UTF-8多字节序列
            has_multibyte = any(b >= 0xC0 for b in raw_bytes)
            if has_multibyte:
                print("  记事本会显示为: UTF-8 (有多字节序列)")
                return "utf-8"
            else:
                print("  记事本会显示为: ANSI (可解码为UTF-8但无多字节序列)")
                return "ansi"
        except UnicodeDecodeError:
            print("  记事本会显示为: ANSI (无法解码为UTF-8)")
            return "ansi"
    
    def analyze_file_content(self, file_path):
        """
        分析文件内容和编码
        """
        print(f"\n=== 分析文件内容 ===")
        
        with open(file_path, 'rb') as f:
            raw_bytes = f.read()
        
        if not raw_bytes:
            print("  文件为空")
            return
        
        print(f"  文件大小: {len(raw_bytes)} 字节")
        print(f"  字节内容(HEX): {raw_bytes.hex(' ')}")
        print(f"  字节内容(ASCII表示): {repr(raw_bytes)}")
        
        # 尝试用不同编码读取
        for encoding in ['ascii', 'gbk', 'utf-8', 'latin-1']:
            try:
                content = raw_bytes.decode(encoding)
                print(f"  用 {encoding} 解码: '{content}'")
            except:
                print(f"  无法用 {encoding} 解码")
    
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
            
            # 3. 创建ANSI编码的version文件（带特殊字节）
            print(f"\n=== 创建version文件 ===")
            version_path = os.path.join(self.selected_folder, "version")
            
            # 如果version文件已存在，先备份
            if os.path.exists(version_path):
                try:
                    backup_path = version_path + ".backup"
                    shutil.copy2(version_path, backup_path)
                    print(f"  已备份原始version文件到: {backup_path}")
                except:
                    print(f"  警告：无法备份version文件")
            
            # 创建带特殊字节的version文件（确保ANSI编码）
            version_created = self.create_ansi_version_file(self.selected_folder, target_version)
            
            if version_created:
                processed_files.append("version")
                
                # 分析创建的文件
                print(f"\n=== 分析创建的version文件 ===")
                self.analyze_file_content(version_path)
                
                # 测试记事本如何识别
                notepad_result_before = self.test_notepad_detection(version_path)
                
                # 4. 清理version文件，删除乱码字符
                print(f"\n=== 清理version文件，删除乱码字符 ===")
                cleaned = self.clean_version_file(self.selected_folder, target_version)
                
                if cleaned:
                    print(f"  ✅ 成功清理version文件")
                    
                    # 分析清理后的文件
                    print(f"\n=== 分析清理后的version文件 ===")
                    self.analyze_file_content(version_path)
                    
                    # 测试记事本如何识别清理后的文件
                    notepad_result_after = self.test_notepad_detection(version_path)
                    
                    # 最终验证
                    print(f"\n=== 最终验证 ===")
                    
                    # 验证文件内容
                    try:
                        with open(version_path, 'r', encoding='gbk') as f:
                            content = f.read().strip()
                        print(f"  用GBK读取version文件内容: '{content}'")
                        
                        # 检查内容是否正确
                        if content == target_version:
                            print(f"  ✅ version文件内容正确")
                        else:
                            print(f"  ⚠️ 警告：version文件内容不匹配")
                            print(f"     期望: '{target_version}'")
                            print(f"     实际: '{content}'")
                            
                            # 尝试用ASCII读取
                            try:
                                with open(version_path, 'rb') as f:
                                    raw_bytes = f.read()
                                ascii_content = ""
                                for b in raw_bytes:
                                    if 32 <= b <= 126:
                                        ascii_content += chr(b)
                                print(f"     提取的ASCII字符: '{ascii_content}'")
                            except:
                                pass
                    except Exception as e:
                        print(f"  读取version文件失败: {e}")
                    
                    print(f"\n记事本识别结果:")
                    print(f"  清理前: {notepad_result_before}")
                    print(f"  清理后: {notepad_result_after}")
                    
                    if notepad_result_after == "ansi":
                        print(f"  ✅ 成功：记事本将version文件识别为ANSI编码")
                    else:
                        print(f"  ⚠️ 注意：记事本可能将version文件识别为{notepad_result_after}")
                else:
                    print(f"\n❌ 清理version文件失败")
            else:
                print(f"\n❌ 创建version文件失败！")
            
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
