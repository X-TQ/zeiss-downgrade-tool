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
            print(f"  转换成功: {os.path.basename(file_path)} -> GBK")
            return True
        except Exception as e:
            print(f"  转换失败 {os.path.basename(file_path)}: {e}")
            return False
    
    def delete_and_recreate_version(self, folder_path, target_version):
        """
        核心解决方案：强制删除旧version文件，并创建新的GBK编码文件。
        """
        print(f"\n=== 处理 version 文件（删除并重建） ===")
        
        # 1. 查找并删除所有名为 `version` 的文件（无论扩展名）
        deleted_files = []
        for filename in os.listdir(folder_path):
            full_path = os.path.join(folder_path, filename)
            if os.path.isfile(full_path):
                name_without_ext = os.path.splitext(filename)[0]
                if name_without_ext.lower() == "version":
                    try:
                        os.remove(full_path)
                        deleted_files.append(filename)
                        print(f"  已删除旧文件: {filename}")
                    except Exception as e:
                        print(f"  删除 {filename} 时出错: {e}")
        
        if deleted_files:
            print(f"  总计删除 {len(deleted_files)} 个旧文件。")
        else:
            print("  未找到名为 'version' 的旧文件。")
        
        # 2. 创建全新的 version 文件（无扩展名），并用 GBK 编码写入版本号
        new_version_path = os.path.join(folder_path, "version")  # 无扩展名
        try:
            # 关键步骤：使用 'gbk' 编码创建和写入文件
            with open(new_version_path, 'w', encoding='gbk') as f:
                f.write(target_version)
            print(f"  已创建新文件 'version'，内容: '{target_version}' (GBK编码)")
            return True
        except Exception as e:
            print(f"  创建新 version 文件失败: {e}")
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
            
            # 3. 【核心修改】删除旧version文件，创建新的version文件
            version_success = self.delete_and_recreate_version(self.selected_folder, target_version)
            
            # 4. 最终验证
            print(f"\n=== 处理结果汇总 ===")
            print(f"已处理普通文件: {', '.join(processed_files) if processed_files else '无'}")
            print(f"Version 文件处理: {'成功' if version_success else '失败'}")
            
            # 最终验证新 version 文件的状态
            final_version_path = os.path.join(self.selected_folder, "version")
            if os.path.exists(final_version_path):
                # 方法一：尝试用GBK读取验证
                try:
                    with open(final_version_path, 'r', encoding='gbk') as f:
                        final_content = f.read().strip()
                    print(f"最终验证 - 文件存在，内容: '{final_content}'")
                    print("状态: 文件可被GBK解码，编码应为ANSI/GBK。")
                except UnicodeDecodeError:
                    print("警告: 无法用GBK解码新文件，编码可能有问题。")
                except Exception as e:
                    print(f"验证时读取文件失败: {e}")
            else:
                print("错误: 最终 version 文件未找到。")
            
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
