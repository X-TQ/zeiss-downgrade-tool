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
    
    def force_ansi_encoding_for_version_file(self, folder_path, target_version):
        """
        强制将version文件保存为ANSI编码，确保Windows记事本识别为ANSI
        """
        print(f"\n=== 强制创建ANSI编码的version文件 ===")
        version_path = os.path.join(folder_path, "version")
        
        try:
            # 方法1：使用二进制方式写入，明确使用GBK编码（Windows中文系统的ANSI）
            print(f"方法1：使用二进制GBK编码写入")
            try:
                gbk_bytes = target_version.encode('gbk')
                with open(version_path, 'wb') as f:
                    f.write(gbk_bytes)
                print(f"  写入成功：{len(gbk_bytes)} 字节")
            except Exception as e:
                print(f"  方法1失败: {e}")
            
            # 方法2：使用Windows-1252编码（西欧ANSI）
            print(f"\n方法2：使用Windows-1252编码写入")
            try:
                with open(version_path, 'w', encoding='cp1252', errors='replace') as f:
                    f.write(target_version)
                print(f"  写入成功")
            except Exception as e:
                print(f"  方法2失败: {e}")
            
            # 验证编码
            self.analyze_file_for_notepad(version_path)
            
            return True
            
        except Exception as e:
            print(f"  创建version文件失败: {e}")
            return False
    
    def analyze_file_for_notepad(self, file_path):
        """
        分析文件，模拟Windows记事本如何识别编码
        """
        print(f"\n--- 记事本编码识别分析 ---")
        
        with open(file_path, 'rb') as f:
            raw_bytes = f.read()
        
        if not raw_bytes:
            print("  文件为空")
            return
        
        print(f"  文件大小: {len(raw_bytes)} 字节")
        print(f"  字节内容(HEX): {raw_bytes.hex(' ', 1)}")
        print(f"  字节内容(ASCII表示): {repr(raw_bytes)}")
        
        # 检查BOM
        if raw_bytes.startswith(b'\xef\xbb\xbf'):
            print("  ⚠️ 检测到UTF-8 BOM - 记事本将显示为UTF-8")
            return "utf-8"
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
            # 进一步检查：是否包含典型的UTF-8多字节序列
            has_utf8_multibyte = any(b >= 0xC0 for b in raw_bytes)
            if has_utf8_multibyte:
                print("  ⚠️ 检测到UTF-8多字节序列 - 记事本可能显示为UTF-8")
                return "utf-8-likely"
            else:
                print("  ✅ 可以解码为UTF-8但无多字节序列 - 记事本可能显示为ANSI")
                return "ansi-likely"
        except UnicodeDecodeError:
            print("  ✅ 无法解码为UTF-8 - 记事本将显示为ANSI")
            return "ansi"
    
    def fix_notepad_ansi_detection(self, file_path, content):
        """
        修复记事本对ANSI编码的识别问题
        确保记事本不会将文件误判为UTF-8
        """
        print(f"\n=== 修复记事本编码识别 ===")
        
        # 方法1：添加一个明确的GBK字符（中文分号）
        print(f"方法1：添加GBK锚定字符")
        try:
            content_with_gbk = content + "；"  # 中文分号，GBK编码为0xA3BB
            with open(file_path, 'wb') as f:
                f.write(content_with_gbk.encode('gbk'))
            print(f"  已添加GBK锚定字符")
            
            # 验证
            with open(file_path, 'rb') as f:
                raw = f.read()
                print(f"  文件字节: {raw.hex(' ', 1)}")
                
                # 检查中文分号的GBK编码
                if b'\xa3\xbb' in raw:
                    print(f"  ✅ 检测到GBK编码的中文分号 (A3 BB)")
                else:
                    print(f"  ❌ 未找到GBK编码的中文分号")
                    
            return True
        except Exception as e:
            print(f"  方法1失败: {e}")
        
        # 方法2：使用Windows-1252编码并添加一个特殊字符
        print(f"\n方法2：使用Windows-1252编码")
        try:
            # 在Windows-1252中，€符号是0x80
            content_with_euro = content + "€"
            with open(file_path, 'wb') as f:
                f.write(content_with_euro.encode('cp1252'))
            print(f"  已添加Windows-1252欧元符号")
            
            # 验证
            with open(file_path, 'rb') as f:
                raw = f.read()
                print(f"  文件字节: {raw.hex(' ', 1)}")
                
                # 检查欧元符号
                if b'\x80' in raw:
                    print(f"  ✅ 检测到Windows-1252欧元符号 (80)")
                else:
                    print(f"  ❌ 未找到Windows-1252欧元符号")
                    
            return True
        except Exception as e:
            print(f"  方法2失败: {e}")
        
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
                except:
                    print(f"  警告：无法备份version文件")
            
            # 使用强制ANSI编码方法
            version_success = self.force_ansi_encoding_for_version_file(self.selected_folder, target_version)
            
            if version_success:
                processed_files.append("version")
                
                # 分析编码识别情况
                print(f"\n=== 最终编码分析 ===")
                detected_encoding = self.analyze_file_for_notepad(version_path)
                
                # 如果记事本可能误判为UTF-8，尝试修复
                if detected_encoding in ["utf-8", "utf-8-likely"]:
                    print(f"\n⚠️ 警告：记事本可能将version文件识别为UTF-8")
                    print(f"正在尝试修复...")
                    
                    if self.fix_notepad_ansi_detection(version_path, target_version):
                        print(f"✅ 已尝试修复记事本编码识别")
                        # 重新分析
                        detected_encoding = self.analyze_file_for_notepad(version_path)
                    else:
                        print(f"❌ 修复失败")
                
                # 最终验证
                print(f"\n=== 最终验证 ===")
                try:
                    with open(version_path, 'r', encoding='gbk') as f:
                        content = f.read().strip()
                    print(f"用GBK读取成功: '{content}'")
                    
                    if content.startswith(target_version):
                        print(f"✅ version文件内容正确")
                    else:
                        print(f"⚠️ 注意：读取的内容以目标版本开头，但可能包含额外字符")
                except Exception as e:
                    print(f"无法用GBK读取: {e}")
                    
                    try:
                        with open(version_path, 'r', encoding='utf-8') as f:
                            content = f.read().strip()
                        print(f"用UTF-8读取成功: '{content}'")
                    except Exception as e2:
                        print(f"也无法用UTF-8读取: {e2}")
                
                print(f"\n记事本识别为: {detected_encoding}")
                
                if detected_encoding in ["ansi", "ansi-likely"]:
                    print(f"✅ 预期结果：记事本应显示为ANSI编码")
                else:
                    print(f"⚠️ 注意：记事本可能显示为{detected_encoding}")
                
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
