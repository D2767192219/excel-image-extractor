#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel图片提取器 - 图形界面版本
基于simple_excel_image_extractor.py开发的GUI工具
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import queue
import logging
import traceback
from datetime import datetime

# 设置日志
log_dir = Path.home() / "Documents" / "ExcelImageExtractor_Logs"
log_dir.mkdir(parents=True, exist_ok=True)
log_file = log_dir / f"excel_extractor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# 导入主要功能模块
try:
    from simple_excel_image_extractor import SimpleExcelImageExtractor
    logging.info("成功导入SimpleExcelImageExtractor")
except Exception as e:
    logging.error(f"导入SimpleExcelImageExtractor失败: {e}")
    logging.error(traceback.format_exc())
    
class RedirectText:
    """用于重定向输出到Text控件"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.queue = queue.Queue()
        self.updating = True
        threading.Thread(target=self._update_text_widget, daemon=True).start()

    def write(self, string):
        self.queue.put(string)
        logging.debug(string.strip())  # 同时记录到日志

    def flush(self):
        pass

    def _update_text_widget(self):
        while self.updating:
            try:
                while True:
                    string = self.queue.get_nowait()
                    self.text_widget.insert(tk.END, string)
                    self.text_widget.see(tk.END)
                    self.text_widget.update()
            except queue.Empty:
                self.text_widget.update()
            except Exception as e:
                logging.error(f"更新文本控件时出错: {e}")
            self.text_widget.after(100, None)

    def stop(self):
        self.updating = False

class ExcelImageExtractorGUI:
    def __init__(self, root):
        try:
            logging.info("开始初始化GUI")
            self.root = root
            self.root.title("Excel图片提取器")
            self.root.geometry("800x600")
            
            # 设置异常处理
            self.root.report_callback_exception = self.handle_exception
            
            # 设置样式
            style = ttk.Style()
            style.configure("TButton", padding=5)
            style.configure("TLabel", padding=5)
            
            # 创建主框架
            self.main_frame = ttk.Frame(root, padding="10")
            self.main_frame.pack(fill=tk.BOTH, expand=True)
            
            # 文件选择区域
            self.create_file_selection_frame()
            
            # 输出区域
            self.create_output_frame()
            
            # 进度条
            self.create_progress_frame()
            
            # 状态变量
            self.processing = False
            
            logging.info("GUI初始化完成")
            
        except Exception as e:
            logging.error(f"GUI初始化失败: {e}")
            logging.error(traceback.format_exc())
            raise

    def handle_exception(self, exc_type, exc_value, exc_traceback):
        """处理未捕获的异常"""
        logging.error("未捕获的异常:", exc_info=(exc_type, exc_value, exc_traceback))
        error_msg = f"发生错误:\n{exc_type.__name__}: {exc_value}"
        messagebox.showerror("错误", error_msg)

    def create_file_selection_frame(self):
        # Excel文件选择
        excel_frame = ttk.LabelFrame(self.main_frame, text="Excel文件", padding="5")
        excel_frame.pack(fill=tk.X, pady=5)
        
        self.excel_path = tk.StringVar()
        ttk.Entry(excel_frame, textvariable=self.excel_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(excel_frame, text="选择文件", command=self.select_excel_file).pack(side=tk.LEFT, padx=5)
        
        # 输出目录选择
        output_frame = ttk.LabelFrame(self.main_frame, text="输出目录", padding="5")
        output_frame.pack(fill=tk.X, pady=5)
        
        self.output_path = tk.StringVar(value="extracted_images")
        ttk.Entry(output_frame, textvariable=self.output_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(output_frame, text="选择目录", command=self.select_output_dir).pack(side=tk.LEFT, padx=5)
        
        # 开始按钮
        self.start_button = ttk.Button(self.main_frame, text="开始提取", command=self.start_extraction)
        self.start_button.pack(pady=10)
        
    def create_output_frame(self):
        # 输出文本区域
        output_frame = ttk.LabelFrame(self.main_frame, text="处理日志", padding="5")
        output_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.output_text = tk.Text(output_frame, height=10, wrap=tk.WORD)
        self.output_text.pack(fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, command=self.output_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_text.configure(yscrollcommand=scrollbar.set)
        
    def create_progress_frame(self):
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(self.main_frame, 
                                      variable=self.progress_var,
                                      maximum=100,
                                      mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=5)
        
    def select_excel_file(self):
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
            
    def select_output_dir(self):
        dirname = filedialog.askdirectory(title="选择输出目录")
        if dirname:
            self.output_path.set(dirname)
            
    def start_extraction(self):
        if self.processing:
            return
            
        try:
            excel_file = self.excel_path.get()
            output_dir = self.output_path.get()
            
            logging.info(f"开始处理文件: {excel_file}")
            logging.info(f"输出目录: {output_dir}")
            
            if not excel_file:
                messagebox.showerror("错误", "请选择Excel文件")
                return
                
            if not os.path.exists(excel_file):
                messagebox.showerror("错误", "Excel文件不存在")
                return
                
            # 禁用按钮
            self.start_button.configure(state='disabled')
            self.processing = True
            
            # 清空输出
            self.output_text.delete(1.0, tk.END)
            
            # 开始进度条
            self.progress.start()
            
            # 重定向输出
            self.redirect = RedirectText(self.output_text)
            sys.stdout = self.redirect
            
            # 在新线程中运行提取过程
            threading.Thread(target=self._run_extraction, args=(excel_file, output_dir), daemon=True).start()
            
        except Exception as e:
            logging.error(f"启动提取过程失败: {e}")
            logging.error(traceback.format_exc())
            self._reset_ui()
            messagebox.showerror("错误", f"启动失败: {str(e)}")

    def _run_extraction(self, excel_file, output_dir):
        try:
            logging.info("开始提取图片")
            extractor = SimpleExcelImageExtractor(excel_file, output_dir)
            extractor.extract_images()
            
            logging.info("提取完成")
            self.root.after(0, self._show_completion_message, output_dir)
            
        except Exception as e:
            logging.error(f"提取过程出错: {e}")
            logging.error(traceback.format_exc())
            self.root.after(0, self._show_error_message, str(e))
            
        finally:
            self.root.after(0, self._reset_ui)

    def _show_completion_message(self, output_dir):
        logging.info(f"处理完成，输出目录: {output_dir}")
        messagebox.showinfo("完成", f"图片提取完成！\n输出目录：{output_dir}")
        
    def _show_error_message(self, error_msg):
        logging.error(f"显示错误消息: {error_msg}")
        messagebox.showerror("错误", f"处理过程中出现错误：\n{error_msg}")
        
    def _reset_ui(self):
        try:
            # 停止进度条
            self.progress.stop()
            
            # 恢复按钮
            self.start_button.configure(state='normal')
            
            # 恢复标准输出
            sys.stdout = sys.__stdout__
            if hasattr(self, 'redirect'):
                self.redirect.stop()
                
            self.processing = False
            logging.info("界面已重置")
            
        except Exception as e:
            logging.error(f"重置界面失败: {e}")
            logging.error(traceback.format_exc())

def main():
    try:
        logging.info("程序启动")
        root = tk.Tk()
        app = ExcelImageExtractorGUI(root)
        logging.info("开始主循环")
        root.mainloop()
    except Exception as e:
        logging.error(f"程序运行失败: {e}")
        logging.error(traceback.format_exc())
        messagebox.showerror("严重错误", f"程序启动失败：\n{str(e)}\n\n详细错误日志已保存到：\n{log_file}")
        sys.exit(1)

if __name__ == "__main__":
    main() 