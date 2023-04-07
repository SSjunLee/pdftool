import sys

import PyPDF2
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import fitz
from PIL import ImageTk, Image


def split_pdf(name, output="", bottom=0, left=0, right=0, top=0):
    if output == "":
        output = name.split(".pdf")[0] + "_out.pdf"

    pdf_file = open(name, 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)

    # 创建一个新的PDF写入器
    pdf_writer = PyPDF2.PdfWriter()

    assert len(pdf_reader.pages) == 1

    page = pdf_reader.pages[0]
    h, w = (page.mediabox.height, page.mediabox.width)
    b = page.mediabox

    assert right <= w and top <= h

    right = w - right
    top = h - top

    # print(b.width, b.height, b.upper_left, b.lower_right, b.upper_right, b.lower_left)
    page.mediabox.lower_left = (left, bottom)
    page.mediabox.upper_right = (right, top)
    # print(b.width, b.height, b.upper_left, b.lower_right, b.upper_right, b.lower_left)

    pdf_writer.add_page(page)
    pdf_output_file = open(output, 'wb')
    pdf_writer.write(pdf_output_file)

    # 关闭文件
    pdf_output_file.close()
    pdf_file.close()


def start_ui():
    root = tk.Tk()
    tk.Tk.report_callback_exception = hook
    app = PDFViewer(master=root)
    app.mainloop()


class PDFViewer(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.pdf_name = ""
        self.start_x = None
        self.start_y = None
        self.end_x = None
        self.end_y = None
        self.rect = None
        self.width = 0
        self.height = 0

        self.label_string = tk.StringVar()
        self.label_string.set("请选择pdf")

        self.create_widgets()

    @property
    def enable(self):
        return self.rect is None and self.pdf_name != ""

    def handleClick(self, event):
        if self.enable:
            self.start_x, self.start_y = (event.x, event.y)
            self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, 1, 1, outline="red")

    def handleMove(self, event):
        if self.enable:
            curX, curY = (event.x, event.y)
            self.canvas.coords(self.rect, self.start_x, self.start_y, curX, curY)

    def handleRealease(self, e):
        if self.enable:
            self.end_x, self.end_y = (e.x, e.y)

    def handle_spit(self, event):
        if not self.enable:
            messagebox.showerror("", "选择一个文件")
            return
        assert self.rect is not None
        left_up = (self.start_x, self.start_y)
        right_down = (self.end_x, self.end_y)

        if left_up > right_down:
            left_up, right_down = right_down, left_up
        try:
            split_pdf(self.pdf_name, left=left_up[0], top=left_up[1], bottom=self.height - right_down[1],
                      right=self.width - right_down[0])
            messagebox.showinfo("提示", "裁剪成功")
        except Exception as ex:
            messagebox.showerror("裁剪错误", ex)

        self.reset()

    def reset(self):
        self.canvas.delete(self.rect)
        self.rect = None
        self.pdf_name = ""
        self.label_string = "请选择pdf"

    def init_btn(self):
        split_btn = tk.Button(self, text="开始切割")
        split_btn.bind("<Button-1>", self.handle_spit)  # #给按钮控件绑定左键单击事件
        split_btn.pack()

        choose_file_btn = tk.Button(self, text="选择文件")
        choose_file_btn.bind("<Button-1>", self.handle_choose_file)
        choose_file_btn.pack()

    def init_label(self):
        self.label = tk.Label(self, textvariable=self.label_string, fg='red', font=('宋体', 10))  # 用于显示警告信息
        self.label.place(x=30, y=120, anchor='nw')
        self.label.pack()

    def handle_choose_file(self, e):
        filename = askopenfilename(title='选择pdf', filetypes=[('pdf文件', '*.pdf')])
        print(filename)
        if filename is not None:
            self.pdf_name = filename
            self.label_string.set(filename)
            self.render_pdf()

    def render_pdf(self):
        if self.pdf_name == "":
            messagebox.showinfo("警告", "请选择一个pdf")
            return
        doc = fitz.open(self.pdf_name)
        print("page count ", doc.page_count)

        page = doc.load_page(0)

        trans = fitz.Matrix(1, 1).prerotate(0)
        pix = page.get_pixmap(matrix=trans, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        self.width = pix.width
        self.height = pix.height

        self.canvas.configure(width=self.width, height=self.height)

        self.photo = ImageTk.PhotoImage(img)
        self.canvas.create_image(0, 0, image=self.photo, anchor='nw')

    def create_widgets(self):
        self.init_btn()
        self.init_label()
        self.canvas = tk.Canvas(self, width=600, height=800)
        self.canvas.pack(fill='both', expand=True)
        self.canvas.bind("<Button-1>", lambda event: self.handleClick(event))
        self.canvas.bind("<ButtonRelease-1>", lambda event: self.handleRealease(event))
        self.canvas.bind("<B1-Motion>", lambda e: self.handleMove(e))


def hook(self, exc_type, exc_value, tb):
    msg = ' Traceback (most recent call last):\n'
    while tb:
        filename = tb.tb_frame.f_code.co_filename
        name = tb.tb_frame.f_code.co_name
        lineno = tb.tb_lineno
        msg += ' File "%.500s", line %d, in %.500s\n' % (filename, lineno, name)
        tb = tb.tb_next

    msg += ' %s: %s\n' % (exc_type.__name__, exc_value)
    print(msg)
    messagebox.showerror(message=msg)


if __name__ == '__main__':
    start_ui()
