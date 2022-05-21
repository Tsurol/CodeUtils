# coding:utf-8
import os
import re
from datetime import datetime

import docx
import xlrd
from docx import Document, ImagePart
from docx.image.image import Image
from docx.oxml import CT_Picture
from docx.text.paragraph import Paragraph


class AnalysisQA(object):
    """解析 excel文件和 word文件中的问答数据。
    :param file_path: 文件路径 list
    情况1：传入一个 excel文件和 一个 word文件，此时 excel文件中问题的答案需要在 word文件中找到，答案可能是图文并茂，也可能是图，也可能是文字。
    情况2：传入一个 excel文件，此时问题和答案都在 excel中，答案是纯文字。
    情况3：传入一个 word文件，此时问题和答案都在 word中，答案可能是图文并茂，也可能是图。
    注：excel文件支持.xlsx和.xls格式；word文件只支持.docx格式。
    """

    def __init__(self, file_path: list):
        self.file_path = file_path
        self.now = datetime.now().strftime("%Y%m%d%H%M%S")
        self.relative_img_path = f'images/{self.now}'
        self.img_root = os.path.join(os.getcwd(), f'{self.relative_img_path}')
        self.docx_path = ''
        self.excel_path = ''
        self.mkdir(self.img_root)
        self.contents = dict()  # 存储word答案中的纯文字
        self.images = dict()  # 存储word答案中的图片

    @staticmethod
    def get_picture(document: Document, paragraph: Paragraph):
        """获取docx文档中的，段落中的图片
        """
        img = paragraph._element.xpath('.//pic:pic')
        if not img:
            return
        img: CT_Picture = img[0]
        embed = img.xpath('.//a:blip/@r:embed')[0]
        related_part: ImagePart = document._part.related_parts[embed]
        image: Image = related_part.image
        return image

    def analysis_excel(self):
        """解析excel文件
        """
        workbook = xlrd.open_workbook(self.excel_path)
        table = workbook.sheets()[0]
        row_num = table.nrows
        col_name = table.row_values(0)
        excel_list = []
        for i in range(1, row_num):
            row = table.row_values(i)
            if row:
                app = {}
                for j in range(len(col_name)):
                    app[col_name[j]] = row[j]
                excel_list.append(app)
        return excel_list

    def analysis_word(self):
        """解析问答word文件，图片存入
        """
        doc = docx.Document(self.docx_path)
        title_list = []
        image_list = []
        content_list = []
        content_str = ''
        for paragraph in doc.paragraphs:
            if "Heading" in paragraph.style.name:
                if title_list:
                    content_list.append(content_str)
                    content_str = ""
                title = paragraph.text
                title_list.append(title)
            if "Normal" in paragraph.style.name:
                content_str += paragraph.text
            image = self.get_picture(doc, paragraph)
            image_list.append(image) if image else None
        else:
            if title_list:
                content_list.append(content_str)
        for title, image, content in zip(title_list, image_list, content_list):
            if os.path.exists('{}/{}.png'.format(self.img_root, title)):
                continue
            with open('{}/{}.png'.format(self.img_root, title), 'wb') as f:
                f.write(image.blob)
            if not self.images.get(title, ""):
                self.images[title] = os.path.join(self.img_root, title) + '.png'
            # 判断字符串是否包含中文，是则表示问题的答案含有文字
            if re.findall('[\u4e00-\u9fa5]+', content):
                if not self.contents.get(title, ""):
                    self.contents[title] = content

    def save_db(self, data):
        """将解析到的问答数据存入数据库，脚本暂不实现
        """
        pass

    @staticmethod
    def mkdir(path):
        """创建多级目录
        :param path: 目录路径
        """
        if not os.path.exists(path):
            os.makedirs(path)

    def get_filepath(self):
        """根据文件后缀名获取文件完整路径
        """
        for file in self.file_path:
            suffix = file.split('.')[-1]
            if suffix in ['docx']:
                self.docx_path = os.path.join(os.getcwd(), file)
            elif suffix in ['xlsx', 'xls']:
                self.excel_path = os.path.join(os.getcwd(), file)
            else:
                raise Exception('暂未支持该类型文件的数据解析')

    def run(self):
        self.get_filepath()
        if len(self.file_path) == 2 and all([self.docx_path, self.excel_path]):
            # 传入了word和excel文件，先解析word再解析excel。
            self.analysis_word()
            self.analysis_excel()
        elif len(self.file_path) == 1 and any([self.docx_path, self.excel_path]):
            if self.excel_path:
                # 只传入了excel文件，解析excel。
                self.analysis_excel()
            else:
                # 只传入了word文件，解析word。
                self.analysis_word()
        else:
            raise Exception('解析异常，请检查文件路径')


if __name__ == '__main__':
    # obj = AnalysisQA(file_path=['docs/test.xlsx'])
    # obj = AnalysisQA(file_path=['docs/知识的力量.docx']).run()
    # obj = AnalysisQA(file_path=['docs/test.xlsx', 'docs/test.docx'])
    pass