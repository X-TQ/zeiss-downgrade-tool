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
            
            with open(file_path, 'w', encoding='gbk', errors='ignore') as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"转换 {os.path.basename(file_path)} 失败: {e}")
            return False

    def ultimate_create_gbk_version_file(self, folder_path, target_version):
        """
        【终极解决方案】创建GBK编码的version文件
        尝试三种不同的方法，确保至少一种成功
        """
        print(f"\n=== 终极方案：创建GBK编码version文件 ===")
        version_path = os.path.join(folder_path, "version")
        
        # 方法1：标准GBK二进制写入（最基本）
        print(f"\n方法1：标准GBK二进制写入")
        try:
            gbk_bytes = target_version.encode('gbk', errors='ignore')
            print(f"  编码 '{target_version}' -> GBK字节: {gbk_bytes.hex(' ')}")
            with open(version_path, 'wb') as f:
                f.write(gbk_bytes)
            print(f"  写入成功")
            if self._analyze_file_encoding(version_path, "方法1"):
                return True
        except Exception as e:
            print(f"  方法1失败: {e}")
        
        # 方法2：使用ASCII兼容的子集（如果版本号只是数字和点）
        print(f"\n方法2：使用纯ASCII写入（GBK兼容）")
        try:
            # 检查目标版本是否只包含ASCII字符（数字和点）
            if all(ord(c) < 128 for c in target_version):
                print(f"  版本 '{target_version}' 是纯ASCII字符")
                with open(version_path, 'wb') as f:
                    f.write(target_version.encode('ascii'))
                print(f"  写入成功")
                if self._analyze_file_encoding(version_path, "方法2"):
                    return True
        except Exception as e:
            print(f"  方法2失败: {e}")
        
        # 方法3：使用Windows ANSI编码别名（cp936）
        print(f"\n方法3：使用Windows代码页936（GBK别名）")
        try:
            with open(version_path, 'w', encoding='cp936', errors='ignore') as f:
                f.write(target_version)
            print(f"  写入成功")
            if self._analyze_file_encoding(version_path, "方法3"):
                return True
        except Exception as e:
            print(f"  方法3失败: {e}")
        
        # 方法4：写入一个明确的非ASCII字符来“锚定”GBK编码
        print(f"\n方法4：添加GBK锚定字符")
        try:
            # 在版本号后添加一个中文分号（GBK明确字符）
            content_with_anchor = target_version + "；"  # 中文分号，不是英文分号
            gbk_bytes = content_with_anchor.encode('gbk')
            print(f"  添加锚定字符 -> 字节: {gbk_bytes.hex(' ')}")
            with open(version_path, 'wb') as f:
                f.write(gbk_bytes)
            print(f"  写入成功")
            if self._analyze_file_encoding(version_path, "方法4"):
                return True
        except Exception as e:
            print(f"  方法4失败: {e}")
        
        return False
    
    def _analyze_file_encoding(self, file_path, method_name):
        """
        深度分析文件编码，提供详细诊断信息
        """
        print(f"\n  [{method_name}] 开始深度编码分析:")
        
        # 1. 读取原始字节
        with open(file_path, 'rb') as f:
            raw_bytes = f.read()
        
        print(f"    文件大小: {len(raw_bytes)} 字节")
        if raw_bytes:
            print(f"    字节内容(HEX): {raw_bytes.hex(' ', 1)}")
            print(f"    字节内容(ASCII表示): {repr(raw_bytes)}")
        
        # 2. 检查BOM（字节顺序标记）
        bom_detected = "无"
        if raw_bytes.startswith(b'\xef\xbb\xbf'):
            bom_detected = "UTF-8 BOM (EF BB BF)"
        elif raw_bytes.startswith(b'\xff\xfe'):
            bom_detected = "UTF-16 LE BOM (FF FE)"
        elif raw_bytes.startswith(b'\xfe\xff'):
            bom_detected = "UTF-16 BE BOM (FE FF)"
        print(f"    BOM检测: {bom_detected}")
        
        # 3. 尝试用不同编码解码
        test_encodings = [
            ('gbk', 'GBK/ANSI'),
            ('utf-8', 'UTF-8'),
            ('utf-8-sig', 'UTF-8 with BOM'),
            ('latin-1', 'Latin-1/ISO-8859-1'),
            ('cp1252', 'Windows-1252')
        ]
        
        successful_encodings = []
        for enc_code, enc_name in test_encodings:
            try:
                # 如果是带BOM的UTF-8，需要特殊处理
                if enc_code == 'utf-8-sig' and bom_detected == "UTF-8 BOM (EF BB BF)":
                    decoded = raw_bytes[3:].decode('utf-8') if len(raw_bytes) > 3 else ""
                else:
                    decoded = raw_bytes.decode(enc_code)
                successful_encodings.append(f"{enc_name}('{decoded}')")
            except UnicodeDecodeError:
                pass
        
        if successful_encodings:
            print(f"    可成功解码的编码: {', '.join(successful_encodings)}")
        else:
            print(f"    警告：无法用任何测试编码解码")
        
        # 4. 判断最可能的编码
        # 优先级：GBK > UTF-8 > 其他
        if any('GBK' in enc for enc in successful_encodings):
            print(f"    【结论】文件很可能已经是GBK/ANSI编码")
            return True
        elif any('UTF-8' in enc for enc in successful_encodings):
            print(f"    【警告】文件似乎是UTF-8编码")
            return False
        else:
            print(f"    【不确定】无法确定编码类型")
            return False
    
    def test_notepad_detection(self, file_path):
        """
        模拟记事本编码检测逻辑
        """
        print(f"\n=== 模拟记事本编码检测 ===")
        
        with open(file_path, 'rb') as f:
            first_1k = f.read(1024)
        
        # 检查BOM
        if first_1k.startswith(b'\xef\xbb\xbf'):
            print("  检测到UTF-8 BOM，记事本会显示为UTF-8")
            return "utf-8"
        
        # 检查是否可能是UTF-8（无BOM）
        try:
            first_1k.decode('utf-8')
            # 进一步检查：UTF-8有效性检查
            # 简单检查：如果所有字节都是ASCII (<128)，记事本可能认为是ANSI
            if all(b < 128 for b in first_1k):
                print("  全部为ASCII字符，记事本可能显示为ANSI")
                return "ansi-ascii"
            else:
                print("  有效的UTF-8序列（无BOM），记事本可能显示为UTF-8")
                return "utf-8"
        except:
            pass
        
        print("  非UTF-8，记事本会显示为ANSI")
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
            
            # 3. 终极方案创建version文件
            version_path = os.path.join(self.selected_folder, "version")
            version_success = self.ultimate_create_gbk_version_file(self.selected_folder, target_version)
            
            if version_success:
                processed_files.append("version")
                
                # 最终验证
                print(f"\n=== 最终验证 ===")
                if os.path.exists(version_path):
                    self._analyze_file_encoding(version_path, "最终检查")
                    self.test_notepad_detection(version_path)
                    
                    # 尝试读取并显示内容
                    try:
                        with open(version_path, 'r', encoding='gbk') as f:
                            content = f.read()
                        print(f"用GBK读取成功: '{content}'")
                    except:
                        try:
                            with open(version_path, 'r', encoding='utf-8') as f:
                                content = f.read()
                            print(f"用UTF-8读取成功: '{content}'")
                        except:
                            print("无法用GBK或UTF-8读取")
            else:
                print(f"\n【严重错误】所有创建version文件的方法都失败了！")
            
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
