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
from xlrd import open_workbook
import docx
import sys


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

    def run(self):
        file_type = self.path.split('.')[-1]
        # print(file_type)
        try:
            if file_type == 'docx':
                context = self.read_docx(self.path)
                # 解析成功后删除文件，把这个功能放到解析里面了
                os.remove(os.getcwd() + "/attach_files/" + self.file_name)
            elif file_type == 'doc':
                context = self.read_doc(self.path)
                # 解析成功后删除文件，把这个功能放到解析里面了
                os.remove(os.getcwd() + "/attach_files/" + self.file_name)
            # elif file_type == 'pptx':
            #     context = self.read_pptx(self.path)
            elif file_type == 'xlsx':
                context = self.read_xlsx(self.path)
                # 解析成功后删除文件，把这个功能放到解析里面了
                os.remove(os.getcwd() + "/attach_files/" + self.file_name)
            elif file_type == 'eml':
                context = self.read_eml(self.path)
                # 解析成功后删除文件，把这个功能放到解析里面了
                os.remove(os.getcwd() + "/attach_files/" + self.file_name)
            else:
                context = ''
            print(context)
            context = self.file_name + "\n" + context + "\n"
            return context
        except Exception as e:
            print('文件读取异常，原因是：', str(e).strip())
            pass


if __name__ == '__main__':
    # spider = ParseAttachFile(r'C:\Users\Bruce Long\Desktop\attach_parse\aaa.docx')
    # spider = ParseAttachFile(r'C:\Users\Bruce Long\Desktop\attach_parse\新建Microsoft PowerPoint 演示文稿.pptx')
    # spider = ParseAttachFile(r'C:\Users\Bruce Long\Desktop\attach_parse\工作通知.doc')
    # spider = ParseAttachFile(r'C:\Users\Bruce Long\Desktop\attach_parse\wx_data.xlsx')
    spider = ParseAttachFile(r'C:\Users\hbsxw\Desktop\163_0424\111.docx')
    print(spider.run())
    # print(os.getcwd())
