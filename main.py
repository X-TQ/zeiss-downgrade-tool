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
    
    def add_invisible_ansi_char_and_convert(self, file_path):
        """
        强制修改文件为ANSI编码：
        1. 读取文件内容
        2. 在末尾添加一个看不见的ANSI字符
        3. 将整个文件编码修改为ANSI (GBK)
        """
        print(f"\n=== 强制修改文件为ANSI编码 ===")
        
        try:
            # 1. 读取原始文件内容
            original_content = ""
            
            # 尝试用常见编码读取
            encodings_to_try = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'latin-1', 'cp1252']
            for enc in encodings_to_try:
                try:
                    with open(file_path, 'r', encoding=enc) as f:
                        original_content = f.read()
                    print(f"  成功用 {enc} 编码读取文件")
                    break
                except:
                    continue
            
            # 如果还是失败，用二进制读取
            if original_content == "":
                print(f"  无法用标准编码读取，尝试二进制读取")
                with open(file_path, 'rb') as f:
                    raw_bytes = f.read()
                
                # 尝试UTF-8解码
                try:
                    original_content = raw_bytes.decode('utf-8')
                except:
                    try:
                        original_content = raw_bytes.decode('gbk')
                    except:
                        # 最后尝试latin-1，它总是能解码
                        original_content = raw_bytes.decode('latin-1', errors='ignore')
            
            print(f"  原始内容: '{original_content}'")
            
            # 2. 移除可能的BOM字符
            if original_content.startswith('\ufeff'):  # UTF-8 BOM字符
                original_content = original_content[1:]
                print(f"  移除了UTF-8 BOM字符")
            
            # 3. 添加看不见的ANSI字符（使用GBK编码中的控制字符）
            # 使用GBK编码的"零宽空格"或"软连字符" - 这些字符在记事本中不可见
            # ANSI字符 0xAD 是"软连字符"，在大多数情况下不可见
            invisible_char = chr(0xAD)  # 软连字符，通常不可见
            
            # 或者使用另一个不可见字符：0x81（在GBK中是控制字符）
            # 但为了更可靠，我们使用软连字符
            content_with_invisible = original_content.strip() + invisible_char
            
            print(f"  添加了不可见的ANSI字符 (0x{ord(invisible_char):02X})")
            
            # 4. 用GBK编码写入文件（ANSI编码）
            with open(file_path, 'wb') as f:
                gbk_bytes = content_with_invisible.encode('gbk', errors='ignore')
                f.write(gbk_bytes)
            
            print(f"  已将文件编码修改为ANSI (GBK)")
            print(f"  写入的字节: {gbk_bytes.hex(' ')}")
            
            # 5. 验证编码
            self.verify_ansi_encoding(file_path)
            
            return True
            
        except Exception as e:
            print(f"  强制修改文件编码失败: {e}")
            return False
    
    def verify_ansi_encoding(self, file_path):
        """
        验证文件是否为ANSI编码
        """
        print(f"\n  验证文件编码:")
        
        try:
            with open(file_path, 'rb') as f:
                raw_bytes = f.read()
            
            print(f"    文件大小: {len(raw_bytes)} 字节")
            print(f"    字节内容(HEX): {raw_bytes.hex(' ', 1)}")
            
            # 检查是否有BOM
            if raw_bytes.startswith(b'\xef\xbb\xbf'):
                print(f"    ❌ 检测到UTF-8 BOM - 文件不是纯ANSI编码")
                return False
            
            # 尝试用GBK读取（ANSI）
            try:
                with open(file_path, 'r', encoding='gbk') as f:
                    content = f.read()
                print(f"    ✅ 可以用GBK编码读取")
                
                # 检查是否包含我们添加的不可见字符
                if chr(0xAD) in content:
                    print(f"    ✅ 检测到不可见ANSI字符 (0xAD)")
                else:
                    print(f"    ⚠️ 未检测到不可见ANSI字符")
                
                return True
            except Exception as e:
                print(f"    ❌ 无法用GBK编码读取: {e}")
                return False
                
        except Exception as e:
            print(f"    验证失败: {e}")
            return False
    
    def clean_version_file_content(self, file_path, target_version):
        """
        清理version文件内容，确保只有版本号，没有乱码字符
        """
        print(f"\n=== 清理version文件内容 ===")
        
        try:
            # 读取文件内容
            with open(file_path, 'rb') as f:
                raw_bytes = f.read()
            
            print(f"  原始字节: {raw_bytes.hex(' ')}")
            
            # 提取所有可打印的ASCII字符（包括数字和小数点）
            clean_bytes = bytearray()
            for b in raw_bytes:
                # ASCII可打印字符（包括数字、小数点、换行符等）
                if (32 <= b <= 126) or b in [10, 13]:  # 10=换行，13=回车
                    clean_bytes.append(b)
            
            if clean_bytes:
                clean_str = clean_bytes.decode('ascii')
                print(f"  提取的ASCII内容: '{clean_str}'")
                
                # 使用正则表达式提取版本号
                import re
                version_pattern = r'\d+\.\d+'
                matches = re.findall(version_pattern, clean_str)
                
                if matches:
                    final_version = matches[0]
                    print(f"  匹配到的版本号: {final_version}")
                    
                    # 如果匹配到的版本号与目标版本号不同，使用目标版本号
                    if final_version != target_version:
                        print(f"  ⚠️ 匹配的版本号与目标版本号不同，使用目标版本号")
                        final_version = target_version
                else:
                    print(f"  ⚠️ 未找到版本号模式，使用目标版本号")
                    final_version = target_version
            else:
                print(f"  ⚠️ 未提取到ASCII内容，使用目标版本号")
                final_version = target_version
            
            # 重新写入文件：版本号 + 不可见ANSI字符，用GBK编码
            invisible_char = chr(0xAD)  # 软连字符
            final_content = final_version + invisible_char
            
            with open(file_path, 'wb') as f:
                gbk_bytes = final_content.encode('gbk')
                f.write(gbk_bytes)
            
            print(f"  最终写入内容: '{final_version}' + 不可见字符")
            print(f"  最终字节: {gbk_bytes.hex(' ')}")
            
            # 验证最终内容
            try:
                with open(file_path, 'r', encoding='gbk') as f:
                    read_content = f.read()
                
                # 移除不可见字符用于显示
                visible_content = read_content.replace(chr(0xAD), '')
                print(f"  读取的内容（移除不可见字符）: '{visible_content}'")
                
                if visible_content == final_version:
                    print(f"  ✅ 文件内容正确")
                    return True
                else:
                    print(f"  ⚠️ 文件内容不匹配: '{visible_content}' != '{final_version}'")
                    return False
                    
            except Exception as e:
                print(f"  验证最终内容失败: {e}")
                return False
                
        except Exception as e:
            print(f"  清理version文件失败: {e}")
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
                # 替换不可见字符为可见表示
                visible_content = content.replace(chr(0xAD), '[0xAD]')
                print(f"  用 {encoding} 解码: '{visible_content}'")
            except:
                print(f"  无法用 {encoding} 解码")
    
    def delete_backup_files(self, folder_path):
        """
        删除备份文件
        """
        print(f"\n=== 删除备份文件 ===")
        backup_path = os.path.join(folder_path, "version.backup")
        
        if os.path.exists(backup_path):
            try:
                os.remove(backup_path)
                print(f"  已删除备份文件: {backup_path}")
                return True
            except Exception as e:
                print(f"  删除备份文件失败: {e}")
                return False
        else:
            print(f"  备份文件不存在，无需删除")
            return True
    
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
            
            # 3. 处理version文件
            print(f"\n=== 处理version文件 ===")
            version_path = os.path.join(self.selected_folder, "version")
            
            # 如果version文件已存在，先备份（后续会删除备份）
            backup_created = False
            if os.path.exists(version_path):
                try:
                    backup_path = version_path + ".backup"
                    shutil.copy2(version_path, backup_path)
                    backup_created = True
                    print(f"  已备份原始version文件到: {backup_path}")
                    
                    # 分析原始文件的编码
                    print(f"\n  分析原始version文件:")
                    self.analyze_file_content(version_path)
                except:
                    print(f"  警告：无法备份version文件")
            
            # 步骤3.1：先创建一个包含版本号的version文件
            print(f"\n  步骤1：创建包含版本号的version文件")
            with open(version_path, 'w', encoding='utf-8') as f:
                f.write(target_version)
            
            # 步骤3.2：强制修改文件为ANSI编码并添加不可见字符
            print(f"\n  步骤2：强制修改为ANSI编码并添加不可见字符")
            ansi_converted = self.add_invisible_ansi_char_and_convert(version_path)
            
            if ansi_converted:
                processed_files.append("version")
                
                # 步骤3.3：清理文件内容，确保只有版本号和不可见字符
                print(f"\n  步骤3：清理文件内容，确保只有版本号和不可见字符")
                cleaned = self.clean_version_file_content(version_path, target_version)
                
                if cleaned:
                    # 步骤3.4：删除备份文件
                    if backup_created:
                        self.delete_backup_files(self.selected_folder)
                    
                    # 最终验证
                    print(f"\n=== 最终验证 ===")
                    
                    # 分析最终文件
                    self.analyze_file_content(version_path)
                    
                    # 测试记事本如何识别
                    notepad_result = self.test_notepad_detection(version_path)
                    
                    # 验证文件内容
                    try:
                        with open(version_path, 'r', encoding='gbk') as f:
                            content = f.read()
                        
                        # 移除不可见字符用于比较
                        visible_content = content.replace(chr(0xAD), '')
                        print(f"\n  用GBK读取version文件内容: '{visible_content}'")
                        
                        # 检查内容是否正确
                        if visible_content == target_version:
                            print(f"  ✅ version文件内容正确")
                        else:
                            print(f"  ⚠️ 警告：version文件内容不匹配")
                            print(f"     期望: '{target_version}'")
                            print(f"     实际: '{visible_content}'")
                    except Exception as e:
                        print(f"  读取version文件失败: {e}")
                    
                    print(f"\n记事本识别结果: {notepad_result}")
                    if notepad_result == "ansi":
                        print(f"  ✅ 成功：记事本将version文件识别为ANSI编码")
                    else:
                        print(f"  ⚠️ 注意：记事本可能将version文件识别为{notepad_result}")
                        
                    # 额外验证：尝试读取并显示（不应看到乱码）
                    print(f"\n  用户视角验证（记事本打开效果）:")
                    with open(version_path, 'rb') as f:
                        raw_bytes = f.read()
                    
                    # 模拟记事本显示：用GBK解码并移除不可见字符
                    try:
                        decoded = raw_bytes.decode('gbk')
                        visible = decoded.replace(chr(0xAD), '')
                        print(f"  用户将看到: '{visible}'")
                        print(f"  文件大小: {len(raw_bytes)} 字节")
                    except:
                        print(f"  无法模拟记事本显示")
                else:
                    print(f"\n❌ 清理version文件失败")
            else:
                print(f"\n❌ 强制修改ANSI编码失败！")
            
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
