"""
送货单对账单生成工具 - GUI版本
"""
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
from pathlib import Path
import sys
import os

# 导入核心功能
from merge_delivery_orders import merge_delivery_orders, create_statement
import pandas as pd


class DeliveryOrderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("送货单对账单生成工具")
        self.root.geometry("700x550")

        # 设置默认路径
        self.raw_data_path = tk.StringVar(value="raw-data")
        self.output_path = tk.StringVar(value="output")

        self.setup_ui()

    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 标题
        title_label = ttk.Label(main_frame, text="送货单对账单生成工具",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)

        # 原始数据路径
        ttk.Label(main_frame, text="原始数据文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.raw_data_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_raw_data).grid(row=1, column=2)

        # 输出路径
        ttk.Label(main_frame, text="输出文件夹:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_output).grid(row=2, column=2)

        # 分隔线
        ttk.Separator(main_frame, orient='horizontal').grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # 操作按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)

        self.run_button = ttk.Button(button_frame, text="生成对账单",
                                     command=self.run_generation, width=20)
        self.run_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="打开输出文件夹",
                  command=self.open_output_folder, width=20).pack(side=tk.LEFT, padx=5)

        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # 日志输出区域
        ttk.Label(main_frame, text="运行日志:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15, width=80)
        self.log_text.grid(row=7, column=0, columnspan=3, pady=5)

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)

    def browse_raw_data(self):
        folder = filedialog.askdirectory(title="选择原始数据文件夹")
        if folder:
            self.raw_data_path.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_path.set(folder)

    def open_output_folder(self):
        output_dir = self.output_path.get()
        if os.path.exists(output_dir):
            if sys.platform == 'darwin':  # macOS
                os.system(f'open "{output_dir}"')
            elif sys.platform == 'win32':  # Windows
                os.startfile(output_dir)
            else:  # Linux
                os.system(f'xdg-open "{output_dir}"')
        else:
            messagebox.showwarning("警告", "输出文件夹不存在")

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def run_generation(self):
        # 在新线程中运行，避免界面冻结
        thread = threading.Thread(target=self._run_generation_thread)
        thread.daemon = True
        thread.start()

    def _run_generation_thread(self):
        try:
            # 禁用运行按钮
            self.run_button.config(state='disabled')
            self.progress.start()

            # 清空日志
            self.log_text.delete(1.0, tk.END)

            raw_data_dir = self.raw_data_path.get()
            output_dir = self.output_path.get()

            self.log("=" * 50)
            self.log("开始处理送货单数据...")
            self.log(f"原始数据文件夹: {raw_data_dir}")
            self.log(f"输出文件夹: {output_dir}")
            self.log("=" * 50)

            # 检查原始数据文件夹是否存在
            if not os.path.exists(raw_data_dir):
                self.log(f"错误: 原始数据文件夹不存在: {raw_data_dir}")
                messagebox.showerror("错误", "原始数据文件夹不存在")
                return

            # 创建输出目录
            Path(output_dir).mkdir(parents=True, exist_ok=True)

            # 合并送货单
            output_file = os.path.join(output_dir, 'merged_delivery_orders.xlsx')

            # 重定向print输出到日志
            import io
            from contextlib import redirect_stdout

            log_stream = io.StringIO()
            with redirect_stdout(log_stream):
                df_summary = merge_delivery_orders(
                    raw_data_dir=raw_data_dir,
                    output_file=output_file
                )

            # 显示合并过程的日志
            for line in log_stream.getvalue().split('\n'):
                if line.strip():
                    self.log(line)

            # 读取详细数据用于生成对账单
            df_all = pd.read_excel(output_file, sheet_name='详细数据')

            # 转换日期列为datetime类型
            df_all['日期'] = pd.to_datetime(df_all['日期'])

            # 提取年月
            df_all['年月'] = df_all['日期'].dt.to_period('M')

            # 按客户和年月分组
            grouped = df_all.groupby(['客户', '年月'])

            self.log(f"\n开始生成对账单...")
            self.log(f"共有 {len(grouped)} 个客户月份组合\n")

            # 创建输出目录
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)

            # 为每个客户的每个月生成对账单
            skipped_count = 0
            generated_count = 0

            for (customer, year_month), group_data in grouped:
                # 创建客户文件夹
                customer_dir = output_path / customer
                customer_dir.mkdir(exist_ok=True)

                # 生成文件名
                statement_file = customer_dir / f'statement_{customer}_{year_month}.xlsx'

                # 检查文件是否已存在
                if statement_file.exists():
                    self.log(f"对账单已存在，跳过: {statement_file.name}")
                    skipped_count += 1
                    continue

                # 格式化年月显示
                year_month_str = f'{year_month.year}年{year_month.month}月'

                # 生成对账单
                log_stream = io.StringIO()
                with redirect_stdout(log_stream):
                    create_statement(
                        group_data,
                        customer_name=customer,
                        year_month=year_month_str,
                        output_file=str(statement_file)
                    )

                # 显示生成日志
                for line in log_stream.getvalue().split('\n'):
                    if line.strip():
                        self.log(line)

                generated_count += 1

            self.log("\n" + "=" * 50)
            self.log("所有对账单生成完成！")
            self.log(f"新生成: {generated_count} 个对账单")
            self.log(f"已跳过: {skipped_count} 个对账单")
            self.log(f"文件保存位置: {output_dir}")
            self.log("=" * 50)

            messagebox.showinfo("完成", f"生成完成！\n新生成: {generated_count} 个对账单\n已跳过: {skipped_count} 个对账单")

        except Exception as e:
            self.log(f"\n错误: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            messagebox.showerror("错误", f"处理过程中出错:\n{str(e)}")

        finally:
            # 恢复按钮和进度条
            self.progress.stop()
            self.run_button.config(state='normal')


def main():
    root = tk.Tk()
    app = DeliveryOrderApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
