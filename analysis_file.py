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
        self.qa_list = list()  # 最终解析出的qa列表

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

    def save_qa(self, excel_data: list, flag: str = ''):
        """将问题和答案组装存入内存中
        情况1：问题和答案都在 excel中, flag='excel'
        情况2：问题在 excel中，答案在 word中, flag='both'
        :param excel_data: excel解析后的问答数据
        :param flag: 标识字段，标识哪种类型的解析
        """
        if flag == 'both':
            for row in excel_data:
                que, ans = self.get_qa(row)
                self.qa_list.append({que: ans})
        elif flag == 'excel':
            for row in excel_data:
                que, ans = row.get('问题', ''), row.get('答案', '')
                self.qa_list.append({que: ans})
        else:
            raise Exception('未知flag')

    def get_qa(self, row):
        """获取 excel中每个问题和对应答案
        :param row: excel中每一行数据
        :return que: 问题, ans: 答案
        """
        question = row.get('问题', '')
        content, img = "", ""
        # excel中的填写的答案，可能是一个word链接
        answer_excel = row.get('答案', '')
        if '.docx' in answer_excel:
            # 存在用户上报的问答中，没有填写序号或者根据序号没找到对应答案的情况，都将序号置为0，item['答案']=''
            serial_mum = row.get('序号', '')
            serial_num_tmp = int(serial_mum) if serial_mum and isinstance(serial_mum, float) else 0
            serial_mum_str = str(serial_num_tmp) + '.'
            for key, value in self.contents.items():
                if not content:
                    content = '<p>{}</p>'.format(value) if key.startswith(serial_mum_str) else content
            for key, value in self.images.items():
                if not img:
                    img = '<img class="global-width" src="/{}/{}.png" alt="{}"/>'.format(
                        self.relative_img_path, key, key) \
                        if key.startswith(serial_mum_str) else img
        answer = content + img
        return question, answer

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
            excel_data = self.analysis_excel()
            self.save_qa(excel_data, flag='both')
        elif len(self.file_path) == 1 and self.excel_path:
            # 只传入了excel文件，解析excel。
            excel_data = self.analysis_excel()
            self.save_qa(excel_data, flag='excel')
        else:
            raise Exception('解析异常，请检查文件路径')
        print(self.qa_list)


if __name__ == '__main__':
    # obj = AnalysisQA(file_path=['docs/问答数据导入模板.xlsx']).run()
    obj = AnalysisQA(file_path=['docs/问答数据导入模板1.xlsx', 'docs/知识的力量.docx']).run()
    pass
