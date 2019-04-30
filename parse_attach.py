# #!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time : 2019/4/24 14:10
# @Author : BruceLong
# @FileName: parse_attach1.py
# @Email   : 18656170559@163.com
# @Software: PyCharm
# @Blog ：http://www.cnblogs.com/yunlongaimeng/
# coding:utf-8
import os

# import win32com.client
import pdfplumber
from xlrd import open_workbook
import docx

from config import FILE_TYPE


class ParseAttachFile(object):
    def __init__(self, path, file_name):
        self.path = path + file_name
        self.file_name = file_name
        pass

    def read_doc(self, path):
        # data = os.system("cd  /data/test/antiword-0.37 && ./antiword ./工作通知.doc")
        try:
            data = os.popen("cd  /data/test/antiword-0.37 && ./antiword {}".format(path)).read()
            context = data.strip().replace("\n", "")
            if context:
                return context
        except:
            return ''

    # def read_pptx(self, path):
    #     try:
    #         note = ''
    #         ppt = win32com.client.Dispatch('PowerPoint.Application')
    #         ppt.Visible = 1
    #         pptSel = ppt.Presentations.Open(path)
    #         slide_count = pptSel.Slides.Count
    #         for i in range(1, slide_count + 1):
    #             shape_count = pptSel.Slides(i).Shapes.Count
    #             for j in range(1, shape_count + 1):
    #                 if pptSel.Slides(i).Shapes(j).HasTextFrame:
    #                     s = pptSel.Slides(i).Shapes(j).TextFrame.TextRange.Text
    #                     note += (' ' + s.strip())
    #         ppt.Quit()
    #         return note.strip()
    #     except:
    #         return ''

    def read_xlsx(self, path):
        try:
            note = ''
            wb = open_workbook(path)
            for s in wb.sheets():
                # print('Sheet:', s.name)
                for row in range(s.nrows):
                    for col in range(s.ncols):
                        value = (s.cell(row, col).value)
                        note += ' ' + str(value)
            note = (note.strip())
            return note.strip()
        except:
            return ''

    def read_docx(self, path):
        try:
            note = ''
            file = docx.Document(path)
            for para in file.paragraphs:
                note += (' ' + para.text.strip())
            note = (note.strip())
            return note
        except:
            return ''

    def read_eml(self, path):
        try:
            with open(path, 'r') as f:
                data = f.read().strip().replace("\n", "")
            return data
        except:
            return ''

    def read_txt(self, path):
        try:
            with open(path, 'r') as f:
                data = f.read().strip().replace('\n', '')
            # print(data)
            return data
        except:
            return ''

    def read_pdf(self, path):
        try:
            data = ''
            with pdfplumber.open(path) as pdf:
                for temp in pdf.pages:
                    data += temp.extract_text().strip().replace('\n', '')
            return data
        except:
            return ''

    def run(self):
        file_type = self.path.split('.')[-1]
        if file_type in FILE_TYPE:
            context = eval('self.read_{}'.format(file_type))(self.path)
            os.remove(os.getcwd() + "/attach_files/" + self.file_name)
        else:
            print('暂时不支持的文件格式')
            context = ''
        context = self.file_name + "\n" + context + "\n"
        print(context)
        return context


if __name__ == '__main__':
    path = r'C:\Users\Bruce Long\Desktop\163_mail\163_0430\attach_files\ '.strip()
    file_name = 'pdf_f5160d266cf44767682cf35b824f47bf_BD卖家后台设置方法指导.pdf'
    spider = ParseAttachFile(path, file_name)
    spider.run()
