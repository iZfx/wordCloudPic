import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import re
import jieba
from datetime import datetime
import os


class WordCloudGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("智能词云生成系统")
        self.root.geometry("1400x900")

        # 初始化变量
        self.df = None
        self.filtered_df = None
        self.current_sheet = tk.StringVar()
        self.selected_column = tk.StringVar()
        self.shape_var = tk.StringVar(value="rectangle")
        self.color_var = tk.StringVar(value="viridis")
        self.width_var = tk.IntVar(value=1920)
        self.height_var = tk.IntVar(value=1080)
        self.radius_var = tk.IntVar(value=400)
        self.word_freq = {}
        self.create_time = ""
        self.wc_image = None  # 新增：存储词云图像对象

        # 设置中文字体
        self.font_path = self.get_font_path()
        if not self.font_path:
            messagebox.showwarning("字体缺失", "未找到中文字体，可能影响中文显示效果")

        # 创建界面
        self.create_widgets()

    def get_font_path(self):
        # 查找系统中可用的中文字体
        font_dirs = [
            "/System/Library/Fonts",  # macOS
            "/usr/share/fonts",  # Linux
            "C:/Windows/Fonts"  # Windows
        ]

        # 常见中文字体名称
        chinese_fonts = [
            "SimHei.ttf",  # 黑体
            "msyh.ttc",  # 微软雅黑
            "STHeiti Medium.ttc",  # 华文黑体
            "FangSong.ttf",  # 仿宋
            "STSong.ttf",  # 宋体
            "SimSun.ttc"  # 宋体
        ]

        for font_dir in font_dirs:
            if os.path.exists(font_dir):
                for font_name in chinese_fonts:
                    font_path = os.path.join(font_dir, font_name)
                    if os.path.exists(font_path):
                        return font_path
        return None

    def create_widgets(self):
        # 主容器使用PanedWindow
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)

        # 左侧数据区
        data_frame = ttk.LabelFrame(main_pane, text="数据管理", width=800)
        main_pane.add(data_frame, weight=1)

        # 文件控制
        file_frame = ttk.Frame(data_frame)
        file_frame.pack(pady=5, fill=tk.X)

        ttk.Button(file_frame, text="导入Excel", command=self.load_file).pack(side=tk.LEFT, padx=5)
        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.pack(side=tk.LEFT, padx=5)

        # Sheet选择
        sheet_frame = ttk.Frame(data_frame)
        sheet_frame.pack(pady=5, fill=tk.X)
        ttk.Label(sheet_frame, text="当前Sheet:").pack(side=tk.LEFT)
        self.sheet_combobox = ttk.Combobox(sheet_frame, textvariable=self.current_sheet, width=20)
        self.sheet_combobox.pack(side=tk.LEFT, padx=5)
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.update_data)

        # 数据表格
        table_frame = ttk.Frame(data_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(table_frame)
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)

        # self.tree.configure(show="headings")  # 隐藏默认的第一列
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set, show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # 右侧控制区
        control_frame = ttk.LabelFrame(main_pane, text="词云设置", width=500)
        main_pane.add(control_frame, weight=0)

        # 列选择
        ttk.Label(control_frame, text="选择分析列:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.column_combobox = ttk.Combobox(control_frame, textvariable=self.selected_column, width=20)
        self.column_combobox.grid(row=0, column=1, padx=5, pady=5)

        # 形状设置
        shape_frame = ttk.LabelFrame(control_frame, text="形状设置")
        shape_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

        ttk.Radiobutton(shape_frame, text="方形", variable=self.shape_var, value="square").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(shape_frame, text="长方形", variable=self.shape_var, value="rectangle").grid(row=0, column=1,
                                                                                                     padx=5)
        ttk.Radiobutton(shape_frame, text="圆形", variable=self.shape_var, value="circle").grid(row=0, column=2, padx=5)

        # 尺寸设置
        size_frame = ttk.LabelFrame(control_frame, text="尺寸设置")
        size_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

        ttk.Label(size_frame, text="宽度:").grid(row=0, column=0)
        ttk.Entry(size_frame, textvariable=self.width_var, width=8).grid(row=0, column=1)
        ttk.Label(size_frame, text="高度:").grid(row=1, column=0)
        ttk.Entry(size_frame, textvariable=self.height_var, width=8).grid(row=1, column=1)
        ttk.Label(size_frame, text="半径:").grid(row=2, column=0)
        ttk.Entry(size_frame, textvariable=self.radius_var, width=8).grid(row=2, column=1)

        # 颜色设置
        color_frame = ttk.LabelFrame(control_frame, text="颜色方案")
        color_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        colors = ["viridis", "plasma", "inferno", "magma", "cividis", "autumn", "winter"]
        ttk.Combobox(color_frame, textvariable=self.color_var, values=colors, width=12).pack(padx=5, pady=5)

        # 生成按钮
        ttk.Button(control_frame, text="生成词云", command=self.generate_wordcloud).grid(row=4, column=0, pady=10)
        # 在控制区添加导出按钮
        ttk.Button(control_frame, text="导出图片", command=self.export_image).grid(row=4, column=1, pady=5)

        # 信息展示区
        info_frame = ttk.LabelFrame(control_frame, text="词云信息")
        info_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

        self.info_text = tk.Text(info_frame, height=10, width=30)
        vsb_info = ttk.Scrollbar(info_frame, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=vsb_info.set)
        self.info_text.pack(side=tk.LEFT, fill=tk.BOTH)
        vsb_info.pack(side=tk.RIGHT, fill=tk.Y)

        # 词云展示区
        self.figure = plt.Figure(figsize=(8, 6))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if file_path:
            self.file_label.config(text=file_path)
            sheets = pd.ExcelFile(file_path).sheet_names
            self.sheet_combobox["values"] = sheets
            if sheets:
                self.current_sheet.set(sheets[0])
                self.update_data()

    def update_data(self, event=None):
        file_path = self.file_label["text"]
        sheet = self.current_sheet.get()
        try:
            self.df = pd.read_excel(file_path, sheet_name=sheet)
            self.filtered_df = self.df.copy()
            self.setup_columns()
            self.show_data()
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")

    def setup_columns(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []
        columns = self.df.columns.tolist()
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col,
                              command=lambda c=col: self.show_filter_menu(c))
            self.tree.column(col, width=120, anchor=tk.W)
        self.column_combobox["values"] = columns
        if columns:
            self.selected_column.set(columns[0])

    def show_data(self):
        self.tree.delete(*self.tree.get_children())
        for _, row in self.filtered_df.iterrows():
            self.tree.insert("", tk.END, values=tuple(row))

    def show_filter_menu(self, column):
        menu = tk.Menu(self.root, tearoff=0)
        try:
            unique_values = self.df[column].dropna().unique()
            unique_values = sorted(unique_values, key=lambda x: str(x))
        except:
            unique_values = []

        current_filter = self.filtered_df[column].unique() if not self.filtered_df.empty else []

        def toggle_filter(value):
            if value in current_filter:
                self.filtered_df = self.filtered_df[self.filtered_df[column] != value]
            else:
                self.filtered_df = pd.concat([self.filtered_df, self.df[self.df[column] == value]])
            self.show_data()

        for value in unique_values[:50]:
            display_value = str(value)[:20]
            menu.add_checkbutton(
                label=display_value,
                command=lambda v=value: toggle_filter(v),
                variable=tk.BooleanVar(value=value in current_filter)
            )

        menu.add_separator()
        menu.add_command(label="重置筛选", command=lambda: self.reset_filter(column))
        menu.post(self.root.winfo_pointerx(), self.root.winfo_pointery())

    def reset_filter(self, column):
        self.filtered_df = self.df.copy()
        self.show_data()

    def clean_text(self, text):
        # 数据清洗：去除非中英文字符
        text = str(text)
        text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', ' ', text)
        # 中文分词
        if any('\u4e00' <= c <= '\u9fff' for c in text):
            return ' '.join(jieba.cut(text))
        return text

    def generate_wordcloud(self):
        if self.filtered_df is None or self.selected_column.get() == "":
            return

        column = self.selected_column.get()
        try:
            # 数据清洗
            text_data = self.filtered_df[column].apply(self.clean_text).tolist()

            # 生成词频
            self.word_freq = {}
            for text in text_data:
                for word in text.split():
                    word = word.strip().lower()
                    if word:
                        self.word_freq[word] = self.word_freq.get(word, 0) + 1

            # 创建mask
            if self.shape_var.get() == "circle":
                size = self.radius_var.get() * 2
                mask = Image.new("L", (size, size), 0)
                draw = ImageDraw.Draw(mask)
                draw.ellipse((0, 0, size, size), fill=255)
                mask = np.array(mask)
                width = height = size
            elif self.shape_var.get() == "rectangle":
                width = self.width_var.get()
                height = self.height_var.get()
                mask = None
            else:  # square
                size = min(self.width_var.get(), self.height_var.get())
                width = height = size
                mask = None

            # 生成词云
            wc = WordCloud(
                font_path=self.font_path,  # 添加字体路径
                width=width,
                height=height,
                background_color="white",
                colormap=self.color_var.get(),
                mask=mask
            ).generate_from_frequencies(self.word_freq)

            # 存储生成的词云图像
            self.wc_image = wc.to_image()

            # 更新信息
            self.create_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.update_info()

            # 显示图像
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.imshow(wc, interpolation="bilinear")
            ax.axis("off")
            self.canvas.draw()

        except Exception as e:
            messagebox.showerror("错误", f"生成失败: {str(e)}")

    def export_image(self):
        if self.wc_image is None:
            messagebox.showwarning("导出失败", "请先生成词云图")
            return

        # 创建信息面板
        info_img = self.create_info_image()

        # 合成图片
        composite = self.composite_images(self.wc_image, info_img)

        # 保存文件
        file_types = [("PNG 图片", "*.png"), ("JPEG 图片", "*.jpg")]
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=file_types,
            title="保存词云图片"
        )

        if file_path:
            composite.save(file_path)
            composite.save(file_path, quality=95)  # 对于JPEG格式，如果需要进一步优化导出图片的质量，可以在保存时添加质量参数
            messagebox.showinfo("保存成功", f"图片已保存至：{file_path}")

    def create_info_image(self):
        # 创建信息图像
        font_size = 16
        line_spacing = 20
        try:
            font = ImageFont.truetype(self.font_path, font_size)
        except:
            font = ImageFont.load_default()

        # 计算图像尺寸
        text_lines = [
            "=== 词云参数 ===",
            f"形状：{self.shape_var.get()}",
            f"尺寸：{self.width_var.get()}x{self.height_var.get()}",
            f"生成时间：{self.create_time}",
            "\n=== 高频词汇 ==="
        ]

        # 添加词频信息
        sorted_words = sorted(self.word_freq.items(), key=lambda x: -x[1])[:50]
        for word, count in sorted_words:
            text_lines.append(f"{word}: {count}次")

        # 计算最大文本宽度
        max_width = max(font.getsize(line)[0] for line in text_lines)
        total_height = len(text_lines) * line_spacing + 20

        # 创建图像
        info_img = Image.new("RGB", (max_width + 50, total_height), "white")
        draw = ImageDraw.Draw(info_img)

        # 绘制文本
        y = 10
        for line in text_lines:
            draw.text((20, y), line, fill="black", font=font)
            y += line_spacing

        return info_img

    def composite_images(self, left_img, right_img):
        # 调整图像高度一致
        max_height = max(left_img.height, right_img.height)
        left_img = left_img.resize((left_img.width, max_height))
        right_img = right_img.resize((right_img.width, max_height))

        # 创建合成图像
        composite = Image.new("RGB", (left_img.width + right_img.width, max_height))
        composite.paste(left_img, (0, 0))
        composite.paste(right_img, (left_img.width, 0))
        return composite
    def update_info(self):
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(tk.END, "词云参数：\n")
        self.info_text.insert(tk.END, f"形状：{self.shape_var.get()}\n")
        self.info_text.insert(tk.END, f"尺寸：{self.width_var.get()}x{self.height_var.get()}\n")
        self.info_text.insert(tk.END, f"生成时间：{self.create_time}\n\n")

        self.info_text.insert(tk.END, "词频统计：\n")
        sorted_words = sorted(self.word_freq.items(), key=lambda x: -x[1])
        for word, count in sorted_words[:50]:  # 显示前50个高频词
            self.info_text.insert(tk.END, f"{word}: {count}次\n")


if __name__ == "__main__":
    root = tk.Tk()
    app = WordCloudGenerator(root)
    root.mainloop()