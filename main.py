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
        """将单个文件内容读取并转换为GBK编码（模仿其他成功文件的处理）"""
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
            # 关键：确保使用 GBK 编码写入，并忽略无法编码的字符
            with open(file_path, 'w', encoding='gbk', errors='ignore') as f:
                f.write(content)
            print(f"  转换成功: {os.path.basename(file_path)} -> GBK")
            return True
        except Exception as e:
            print(f"  转换失败 {os.path.basename(file_path)}: {e}")
            return False
    
    def process_version_file(self, folder_path, target_version):
        """
        独立处理 version 文件：
        1. 查找文件 (支持无扩展名或任意扩展名)
        2. 如果找到，读取并转换编码后写入新版本
        3. 如果没找到，创建新的 version 文件
        """
        print(f"\n=== 开始处理 version 文件 ===")
        version_file_path = None
        
        # 1. 查找 version 文件
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                name_without_ext = os.path.splitext(filename)[0]
                if name_without_ext.lower() == "version":
                    version_file_path = file_path
                    print(f"找到文件: {filename}")
                    break
        
        # 2. 如果找到文件，先转换编码再写入新版本
        if version_file_path and os.path.exists(version_file_path):
            print(f"处理现有文件: {os.path.basename(version_file_path)}")
            # 先备份原内容（如果需要）
            try:
                with open(version_file_path, 'r', encoding='gbk') as f:
                    old_content = f.read().strip()
                    print(f"  原内容: '{old_content}'")
            except:
                # 如果无法用GBK读取，尝试其他编码
                try:
                    with open(version_file_path, 'r', encoding='utf-8') as f:
                        old_content = f.read().strip()
                        print(f"  原内容(UTF-8): '{old_content}'")
                except:
                    old_content = "未知"
            
            # 转换编码并写入新版本号
            success = self.convert_file_to_gbk(version_file_path)
            if success:
                # 重新打开并写入目标版本号
                with open(version_file_path, 'w', encoding='gbk') as f:
                    f.write(target_version)
                print(f"  已写入新版本: '{target_version}' (GBK编码)")
                return True
            else:
                print("  警告：编码转换失败，尝试直接写入...")
                # 转换失败，尝试直接写入
                try:
                    with open(version_file_path, 'w', encoding='gbk') as f:
                        f.write(target_version)
                    print(f"  已直接写入新版本: '{target_version}'")
                    return True
                except Exception as e:
                    print(f"  直接写入失败: {e}")
                    return False
        
        # 3. 如果没找到，创建新的 version 文件
        else:
            print("未找到 version 文件，创建新文件。")
            new_version_path = os.path.join(folder_path, "version")
            try:
                with open(new_version_path, 'w', encoding='gbk') as f:
                    f.write(target_version)
                print(f"已创建新文件 'version'，内容: '{target_version}' (GBK编码)")
                return True
            except Exception as e:
                print(f"创建 version 文件失败: {e}")
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
            
            # 2. 处理 inspset, inspection, username 文件（保持原逻辑）
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
                            if self.convert_file_to_gbk(full_path):
                                processed_files.append(base_name)
                            file_found = True
                            break
                if not file_found:
                    print(f"未找到文件: {base_name}")
            
            # 3. 独立处理 version 文件
            version_success = self.process_version_file(self.selected_folder, target_version)
            
            # 4. 最终验证
            print(f"\n=== 处理结果汇总 ===")
            print(f"已处理普通文件: {', '.join(processed_files) if processed_files else '无'}")
            print(f"Version 文件处理: {'成功' if version_success else '失败'}")
            
            # 验证 version 文件最终状态
            version_check_path = os.path.join(self.selected_folder, "version")
            if os.path.exists(version_check_path):
                try:
                    # 尝试用GBK读取验证
                    with open(version_check_path, 'r', encoding='gbk') as f:
                        final_content = f.read().strip()
                    print(f"最终验证 - Version 文件内容: '{final_content}' (可被GBK解码)")
                    # 尝试用UTF-8读取，如果失败则说明不是UTF-8
                    try:
                        with open(version_check_path, 'r', encoding='utf-8') as f:
                            f.read()
                        print("警告：文件可能仍兼容UTF-8编码，但GBK写入已成功。")
                    except:
                        print("确认：文件不是UTF-8编码。")
                except Exception as e:
                    print(f"验证时读取version文件失败: {e}")
            
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
