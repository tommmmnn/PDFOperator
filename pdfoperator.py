import tkinter
import windnd
from tkinter.messagebox import showinfo, showerror
import PyPDF2
import os
import sys

class PDFOP:

    def __init__(self) -> None:
        self.tk = tkinter.Tk()
        self.tk.geometry('820x480')
        windnd.hook_dropfiles(self.tk, func = self.dragged_files)

        self.file_lists = []
        self.input_files_lable = tkinter.Label(self.tk, text='输入文件序列')
        self.input_files_lable.place(relx=0.01, rely=0.01, relwidth=0.618, relheight=0.05)
        self.output_name = tkinter.Listbox(self.tk)
        self.output_name.place(relx=0.01, rely=0.06, relwidth=0.618, relheight=0.93)

        self.filename_label = tkinter.Label(self.tk, text='输出文件名称')
        self.filename_label.place(relx=0.65, rely=0.01, relwidth=0.349, relheight=0.05)
        self.outputfilename = tkinter.Entry(self.tk)
        self.outputfilename.insert(0, 'output.pdf')
        self.outputfilename.place(relx=0.65, rely=0.07, relwidth=0.349, relheight=0.1)
        self.merge_button = tkinter.Button(self.tk, text = '合并', command=self.merge, bg='#bbdac6')
        self.merge_button.place(relx=0.65, rely=0.25, relwidth=0.349, relheight=0.08)


        self.slice_lable = tkinter.Label(self.tk, text='选择页码(格式为  x-y)')
        self.slice_lable.place(relx=0.65, rely=0.4, relwidth=0.349, relheight=0.05)
        self.slice_range = tkinter.Entry(self.tk)
        self.slice_range.insert(0, '1-3')
        self.slice_range.place(relx=0.65, rely=0.46, relwidth=0.349, relheight=0.1)
        self.slice_button = tkinter.Button(self.tk, text="拆分", command=self.slice, bg='#ede1d1')
        self.slice_button.place(relx=0.65, rely=0.58, relwidth=0.349, relheight=0.08)

    def merge(self):
        merger = PyPDF2.PdfMerger()

        for pdf in self.file_lists:
            merger.append(pdf)
        file_path = os.path.dirname(os.path.realpath(sys.executable))
        file_name = file_path + os.sep + self.outputfilename.get()
        merger.write(file_name)
        showinfo(message=f'输出结果{file_name}')

    def slice(self):
        if self.file_lists == []:
            showerror(message='当前窗口列表中无文件')
            return
        pdf = PyPDF2.PdfReader(open(self.file_lists[0], "rb"))
        pdf_writer = PyPDF2.PdfWriter()

        page_range_l, page_range_r = self.slice_range.get().split('-')
        for page_num in range(int(page_range_l)-1, int(page_range_r)):
            pdf_writer.add_page(pdf.pages[page_num])

        file_path = os.path.dirname(os.path.realpath(sys.executable))
        file_name = file_path + os.sep + self.outputfilename.get()
        with open(file_name, "wb") as f:
            pdf_writer.write(f)
            showinfo(message=f'输出结果{file_name}')

    def dragged_files(self, files):
        for file in files:
            if file not in self.file_lists:
                self.file_lists.append(file.decode('gbk'))
                self.output_name.insert(len(self.file_lists)-1, file.decode('gbk'))


    def run(self):
        self.tk.mainloop()


if __name__ == "__main__":
    mypdf = PDFOP()
    mypdf.run()


    

