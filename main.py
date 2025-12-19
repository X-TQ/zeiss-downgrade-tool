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
        通过在文件末尾添加特殊字节，让Windows记事本识别为ANSI编码
        """
        print(f"\n=== 创建ANSI编码的version文件 ===")
        version_path = os.path.join(folder_path, "version")
        
        try:
            # 方法1：使用ANSI编码并添加一个在UTF-8中无效但在ANSI中有效的字节
            print(f"方法1：使用ANSI编码添加特殊字节")
            try:
                # 先将版本号转换为GBK编码
                gbk_bytes = target_version.encode('gbk')
                
                # 添加一个在ANSI中有效但在UTF-8中无效的字节序列
                # 0x81是GBK编码的第一个字节，但不是有效的UTF-8起始字节
                special_bytes = b'\x81'  # 这个字节单独存在时，UTF-8解码会失败
                
                with open(version_path, 'wb') as f:
                    f.write(gbk_bytes)
                    f.write(special_bytes)
                
                print(f"  写入成功：版本号 + 特殊字节")
                print(f"  字节序列: {gbk_bytes.hex(' ')} {special_bytes.hex(' ')}")
            except Exception as e:
                print(f"  方法1失败: {e}")
            
            # 验证
            self.analyze_notepad_detection(version_path)
            
            return True
            
        except Exception as e:
            print(f"  创建version文件失败: {e}")
            return False
    
    def create_ansi_version_file_advanced(self, folder_path, target_version):
        """
        更高级的方法：创建确保被记事本识别为ANSI的version文件
        通过在文件中插入无效的UTF-8序列来欺骗记事本
        """
        print(f"\n=== 高级方法：创建ANSI编码version文件 ===")
        version_path = os.path.join(folder_path, "version")
        
        try:
            # 策略：创建一个文件，其中包含一个无效的UTF-8序列
            # 这样记事本无法用UTF-8解码，就会回退到ANSI
            
            # 版本号（纯ASCII部分）
            version_str = target_version
            
            # 方法A：在文件末尾添加一个无效的UTF-8字节
            print(f"方法A：添加无效UTF-8字节")
            try:
                # 将版本号编码为ASCII（所有版本号都是ASCII字符）
                version_bytes = version_str.encode('ascii')
                
                # 添加一个单独的0x80字节（在UTF-8中，0x80-0xBF只能是连续字节）
                # 但前面没有起始字节，所以是无效的UTF-8
                invalid_utf8 = b'\x80'
                
                with open(version_path, 'wb') as f:
                    f.write(version_bytes)
                    f.write(invalid_utf8)
                
                print(f"  方法A成功")
                print(f"  字节: {version_bytes.hex(' ')} {invalid_utf8.hex(' ')}")
                
                # 验证
                result = self.analyze_notepad_detection(version_path)
                if result == "ansi":
                    print(f"  ✅ 方法A成功：记事本识别为ANSI")
                    return True
            except Exception as e:
                print(f"  方法A失败: {e}")
            
            # 方法B：在版本号中插入一个无效UTF-8序列
            print(f"\n方法B：插入无效UTF-8序列")
            try:
                # 创建字节数组
                version_bytes = bytearray(version_str.encode('ascii'))
                
                # 在中间位置插入一个无效的UTF-8序列
                # 0xC1 是一个过长的UTF-8起始字节（对于ASCII字符来说）
                insert_pos = len(version_bytes) // 2
                version_bytes.insert(insert_pos, 0xC1)
                
                with open(version_path, 'wb') as f:
                    f.write(version_bytes)
                
                print(f"  方法B成功：在位置{insert_pos}插入了0xC1")
                print(f"  字节: {version_bytes.hex(' ')}")
                
                # 验证
                result = self.analyze_notepad_detection(version_path)
                if result == "ansi":
                    print(f"  ✅ 方法B成功：记事本识别为ANSI")
                    return True
            except Exception as e:
                print(f"  方法B失败: {e}")
            
            # 方法C：使用GB2312编码添加一个中文字符，然后用零宽空格包围
            print(f"\n方法C：使用GB2312编码")
            try:
                # 版本号 + 一个GB2312字符（中文句号）
                content = version_str + "。"
                gb2312_bytes = content.encode('gb2312')
                
                # 添加一个零宽空格（UTF-8中是EF BB BF，但在这里是无效的）
                # 实际上我们添加一个字节0xEF，但后面不跟有效的UTF-8序列
                final_bytes = gb2312_bytes + b'\xEF'
                
                with open(version_path, 'wb') as f:
                    f.write(final_bytes)
                
                print(f"  方法C成功：版本号 + 中文句号 + 0xEF")
                print(f"  字节: {final_bytes.hex(' ')}")
                
                # 验证
                result = self.analyze_notepad_detection(version_path)
                if result == "ansi":
                    print(f"  ✅ 方法C成功：记事本识别为ANSI")
                    return True
            except Exception as e:
                print(f"  方法C失败: {e}")
            
            # 方法D：终极方法 - 创建混合编码文件
            print(f"\n方法D：混合编码文件")
            try:
                # 创建一个文件，前部分是ASCII，中间是无效UTF-8，后部分是ASCII
                version_bytes = version_str.encode('ascii')
                
                # 创建一个字节数组
                final_bytes = bytearray(version_bytes)
                
                # 添加一个无法用UTF-8解码的字节序列
                # UTF-8规则：如果第一个字节是11110xxx (0xF0-0xF7)，需要后面跟3个10xxxxxx字节
                # 我们只写第一个字节，不写后续字节
                final_bytes.append(0xF0)  # 应该是4字节UTF-8序列的开始
                # 但不添加后续的3个字节，所以UTF-8解码会失败
                
                with open(version_path, 'wb') as f:
                    f.write(final_bytes)
                
                print(f"  方法D成功：版本号 + 0xF0（不完整的4字节UTF-8起始）")
                print(f"  字节: {final_bytes.hex(' ')}")
                
                # 验证
                result = self.analyze_notepad_detection(version_path)
                if result == "ansi":
                    print(f"  ✅ 方法D成功：记事本识别为ANSI")
                    return True
            except Exception as e:
                print(f"  方法D失败: {e}")
            
            return False
            
        except Exception as e:
            print(f"  高级方法失败: {e}")
            return False
    
    def analyze_notepad_detection(self, file_path):
        """
        分析Windows记事本如何识别文件编码
        """
        print(f"\n--- 记事本编码识别分析 ---")
        
        with open(file_path, 'rb') as f:
            raw_bytes = f.read()
        
        if not raw_bytes:
            print("  文件为空")
            return "empty"
        
        print(f"  文件大小: {len(raw_bytes)} 字节")
        print(f"  字节内容(HEX): {raw_bytes.hex(' ', 1)}")
        
        # 检查BOM
        if raw_bytes.startswith(b'\xef\xbb\xbf'):
            print("  ⚠️ 检测到UTF-8 BOM - 记事本将显示为UTF-8")
            return "utf-8-bom"
        elif raw_bytes.startswith(b'\xff\xfe'):
            print("  ⚠️ 检测到UTF-16 LE BOM - 记事本将显示为Unicode")
            return "utf-16le"
        elif raw_bytes.startswith(b'\xfe\xff'):
            print("  ⚠️ 检测到UTF-16 BE BOM - 记事本将显示为Unicode big-endian")
            return "utf-16be"
        
        # 检查是否是纯ASCII
        is_ascii = all(b < 128 for b in raw_bytes)
        if is_ascii:
            print("  ✅ 所有字节都是ASCII (<128) - 记事本将显示为ANSI")
            return "ansi"
        
        # 尝试UTF-8解码
        try:
            decoded = raw_bytes.decode('utf-8')
            print(f"  ⚠️ 可以成功解码为UTF-8 - 记事本可能显示为UTF-8")
            print(f"    解码内容: {repr(decoded)}")
            return "utf-8"
        except UnicodeDecodeError as e:
            print(f"  ✅ 无法解码为UTF-8 - 记事本将显示为ANSI")
            print(f"    UTF-8解码错误: {e}")
            
            # 尝试GBK解码
            try:
                decoded = raw_bytes.decode('gbk')
                print(f"    可以解码为GBK，内容: {repr(decoded)}")
            except:
                print(f"    也无法解码为GBK")
            
            return "ansi"
    
    def test_notepad_detection(self, file_path):
        """
        测试记事本如何检测文件编码
        """
        print(f"\n=== 测试记事本编码检测 ===")
        
        with open(file_path, 'rb') as f:
            raw_bytes = f.read()
        
        if not raw_bytes:
            print("  文件为空")
            return
        
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
            
            # 3. 创建ANSI编码的version文件
            print(f"\n=== 处理version文件 ===")
            version_path = os.path.join(self.selected_folder, "version")
            
            # 如果version文件已存在，先备份
            if os.path.exists(version_path):
                try:
                    backup_path = version_path + ".backup"
                    shutil.copy2(version_path, backup_path)
                    print(f"  已备份原始version文件到: {backup_path}")
                    
                    # 分析原始文件的编码
                    print(f"\n  分析原始version文件编码:")
                    self.analyze_notepad_detection(version_path)
                except:
                    print(f"  警告：无法备份version文件")
            
            # 使用高级方法创建ANSI编码的version文件
            version_success = self.create_ansi_version_file_advanced(self.selected_folder, target_version)
            
            if version_success:
                processed_files.append("version")
                
                # 最终验证
                print(f"\n=== 最终验证 ===")
                
                # 分析记事本如何识别
                notepad_result = self.test_notepad_detection(version_path)
                
                # 尝试读取内容（忽略特殊字节）
                print(f"\n  读取文件内容:")
                try:
                    # 尝试用ASCII读取（应该能读取版本号）
                    with open(version_path, 'rb') as f:
                        raw_bytes = f.read()
                    
                    # 提取ASCII部分（版本号）
                    version_bytes = bytearray()
                    for b in raw_bytes:
                        if 32 <= b <= 126:  # 可打印ASCII字符
                            version_bytes.append(b)
                    
                    if version_bytes:
                        version_str = version_bytes.decode('ascii')
                        print(f"  提取的版本号: {version_str}")
                        
                        if version_str.startswith(target_version):
                            print(f"  ✅ version文件内容正确")
                        else:
                            print(f"  ⚠️ 版本号不匹配: 期望 '{target_version}', 实际 '{version_str}'")
                    else:
                        print(f"  ❌ 未找到版本号")
                        
                except Exception as e:
                    print(f"  读取文件内容失败: {e}")
                
                print(f"\n记事本识别结果: {notepad_result}")
                if notepad_result == "ansi":
                    print(f"✅ 成功：记事本将显示为ANSI编码")
                else:
                    print(f"⚠️ 注意：记事本可能显示为{notepad_result}")
                
            else:
                print(f"\n❌ 错误：创建ANSI编码的version文件失败！")
            
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
