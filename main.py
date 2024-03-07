import random
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import openpyxl


def center_window(root: tk.Tk, w, h):
    # 获取屏幕 宽、高
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    # 计算 x, y 位置
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))


def deal_data(filepath):
    df = pd.read_excel(filepath)
    columns = df.columns.values.tolist()
    assert "序号" in columns, "需要一个名为 “序号” 的列表！"
    assert "姓名" in columns, "需要一个名为 “姓名” 的列表！"
    return [f"{row['序号']} {row['姓名']}" for idx, row in df.iterrows()]


class UploadWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("上传")
        center_window(self, 350, 100)

        self.label_filepath = tk.Label(self, text="文件路径")
        self.label_filepath.grid(row=0, column=0, padx=(10, 0), pady=10)
        self.entry_filepath = tk.Entry(self)
        self.entry_filepath.grid(row=0, column=1, columnspan=2, padx=(0, 10), ipadx=60)

        self.btn_upload_file = tk.Button(self, text="上传文件", command=self.upload_file)
        self.btn_upload_file.grid(row=2, column=1, pady=10, ipadx=30)
        self.btn_parse_data = tk.Button(self, text="解析数据", command=self.parse_data)
        self.btn_parse_data.grid(row=2, column=2, ipadx=30)

    def upload_file(self):
        filepath = filedialog.askopenfilename(title="请选择一个文件", filetypes=[("Excel", ".xls .xlsx")])
        self.entry_filepath.delete(0, tk.END)
        self.entry_filepath.insert(0, filepath)

    def parse_data(self):
        try:
            data = deal_data(self.entry_filepath.get())
            self.destroy()
            CallWindow(data)
        except Exception as e:
            from tkinter.messagebox import showwarning
            showwarning("警告", f"解析数据失败！\n{e}")


class CallWindow(tk.Tk):
    def __init__(self, data):
        super().__init__()
        self.data = data
        self.running = False

        self.title("有奖问答")
        self.geometry('405x300')
        self.configure(bg='#ECf5FF')

        self.var = tk.StringVar(value="即 将 开 始")
        self.show_label = tk.Label(
            self,
            textvariable=self.var,
            justify='center',
            width=20, height=5,
            bg='#BFEFFF',
            font=('楷体', 40, 'bold'),
            foreground='black'
        )
        self.show_label.pack(pady=20)

        self.btn_start = tk.Button(
            self,
            text='开始',
            command=lambda: self.lottery_start(self.var),
            width=14, height=2,
            bg='#A8A8A8',
            font=('宋体', 18, 'bold')
        )
        self.btn_start.pack(side=tk.LEFT, padx=50)

        self.btn_end = tk.Button(
            self,
            text='结束',
            command=lambda: self.lottery_end(),
            width=14, height=2,
            bg='#A8A8A8',
            font=('宋体', 18, 'bold')
        )
        self.btn_end.pack(side=tk.RIGHT, padx=50)

    def lottery_roll(self, string):
        string.set(random.choice(self.data))
        if self.running:
            self.after(50, self.lottery_roll, string)

    def lottery_start(self, string):
        if self.running:
            return
        self.running = True
        self.lottery_roll(string)

    def lottery_end(self):
        if self.running:
            self.running = False


window = UploadWindow()
window.mainloop()

