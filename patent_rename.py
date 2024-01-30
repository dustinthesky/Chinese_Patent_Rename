import os
import re
import io
import sys
import shutil
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog
from PyQt5.QtGui import QFont

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.folderPath = ""
        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 400, 200)
        self.setWindowTitle('中国专利重命名')

        # 添加程序名称标签
        nameLabel = QLabel('对中国专利重命名', self)
        nameLabel.setFont(QFont('Arial',14))
        nameLabel.move(130, 40)

        # 添加选择文件夹按钮
        folderButton = QPushButton('选择文件夹', self)
        folderButton.setFont(QFont('Arial',14))
        folderButton.move(150, 100)
        folderButton.clicked.connect(self.showFolderDialog)

    def showFolderDialog(self):
        self.folderPath = QFileDialog.getExistingDirectory(self, '选择文件夹')
        if self.folderPath:
            self.processPDFs()
        QApplication.quit()
        # print('选择的文件夹路径:', folderPath)

    def pdf_text_extractor(self, path):
        with open(path, 'rb') as fh:
            rsrcmgr = PDFResourceManager()
            retstr = io.StringIO()
            device = TextConverter(rsrcmgr, retstr, laparams=None)
            interpreter = PDFPageInterpreter(rsrcmgr, device)

            for page in PDFPage.get_pages(fh, caching=True, check_extractable=True):
                interpreter.process_page(page)

            text = retstr.getvalue()

        device.close()
        retstr.close()
        return text

    def move_pdf_to_duplicated_folder(self, pdf_path):
        # 检查"Duplicated"文件夹是否存在
        duplicated_folder = 'Duplicated'
        if not os.path.exists(duplicated_folder):
            # 如果不存在，创建文件夹
            os.makedirs(duplicated_folder)
        
        # 构建目标文件的路径
        destination_path = os.path.join(duplicated_folder, os.path.basename(pdf_path))
        
        # 移动文件
        try:
            shutil.move(pdf_path, destination_path)
            print(f"File '{pdf_path}' moved to '{destination_path}' successfully.")
        except IOError as e:
            print(f"Unable to move file. {e}")
        except Exception as e:
            print(f"Unexpected error: {e}")

    def extract_patent_info(self, text):
        info = {}
        # 判断专利类型
        patent_type_match = re.search(r"\(12\)(发明专利申请|发明专利|实用新型专利|外观设计专利)", text)
        patent_type = "Unknown"

        if patent_type_match:
            if patent_type_match.group(1) == "发明专利申请":
                patent_type = "发明专利申请"
                # 提取发明专利相关信息
                # info['申请公布号'], info['申请公布日'] = re.findall(r'\bCN \d+ A\b', text)
                number_date_match = re.search(r'(CN \d+ [A-Z])\s*(\d{4}.\d{2}.\d{2})', text)
                if number_date_match:
                    info['申请公布号'] = number_date_match.group(1).replace(' ', '') 
                    info['申请公布日'] = number_date_match.group(2).replace('.', '-')             
                # 使用正则表达式提取申请人信息
                # applicants = re.findall(r"申请人\s+((?:(?!地址).)*?)(?:地址|\s+\(\d{2}\)|$)", text)
                applicants = re.findall(r"申请人\s+((?:(?!\(\d{2}\)).)*?)(?:地址|(?=\(\d{2}\))|$)", text)

                # 移除列表中每个字符串末尾可能存在的空白字符，并将内部空格替换为破折号
                # applicants = [applicant.rstrip().replace(" ", "-") for applicant in applicants]
                applicants = [re.sub(r"\s+", "-", applicant.strip())for applicant in applicants]
                # 格式化输出，连续字符之间用 '-' 连接起来
                applicants = '-'.join(applicants)

                invention_name_match = re.search(r"\(54\)发明名称(.*?)\(57\)", text, re.S)
                
                info['申请人'] = applicants if applicants else "Unknown"
                info['发明名称'] = invention_name_match.group(1).strip() if invention_name_match else "Unknown"
            elif patent_type_match.group(1) == "发明专利":
                patent_type = "发明专利"
                # 提取实用新型专利相关信息
                # info['授权公告号'], info['授权公告日'] = re.findall(r'\bCN \d+ A\b', text)
                number_date_match = re.search(r'(CN \d+ [A-Z])\s*(\d{4}.\d{2}.\d{2})', text)
                if number_date_match:
                    info['授权公布号'] = number_date_match.group(1).replace(' ', '') 
                    info['授权公布日'] = number_date_match.group(2).replace('.', '-')   
                                # 使用正则表达式提取申请人信息
                # applicants = re.findall(r"专利权人\s+((?:(?!地址).)*?)(?:地址|\s+\(\d{2}\)|$)", text)
                applicants = re.findall(r"专利权人\s+((?:(?!\(\d{2}\)).)*?)(?:地址|(?=\(\d{2}\))|$)", text)
                # 移除列表中每个字符串末尾可能存在的空白字符，并将内部空格替换为破折号
                # applicants = [applicant.rstrip().replace(" ", "-") for applicant in applicants]
                applicants = [re.sub(r"\s+", "-", applicant.strip())for applicant in applicants]
                # 格式化输出，连续字符之间用 '-' 连接起来
                applicants = '-'.join(applicants)

                utility_model_name_match = re.search(r"\(54\)发明名称(.*?)\(57\)", text, re.S)
                
                info['专利权人'] = applicants if applicants else "Unknown"
                info['发明名称'] = utility_model_name_match.group(1).strip() if utility_model_name_match else "Unknown"
            elif patent_type_match.group(1) == "外观设计专利":
                patent_type = "外观设计专利"
                # 提取实用新型专利相关信息
                # info['授权公告号'], info['授权公告日'] = re.findall(r'\bCN \d+ A\b', text)
                number_date_match = re.search(r'(CN \d+ [A-Z])\s*(\d{4}.\d{2}.\d{2})', text)
                if number_date_match:
                    info['授权公布号'] = number_date_match.group(1).replace(' ', '') 
                    info['授权公布日'] = number_date_match.group(2).replace('.', '-')  
                # applicants = re.findall(r"专利权人\s+((?:(?!地址).)*?)(?:地址|\s+\(\d{2}\)|$)", text)
                applicants = re.findall(r"专利权人\s+((?:(?!\(\d{2}\)).)*?)(?:地址|(?=\(\d{2}\))|$)", text)
                # 移除列表中每个字符串末尾可能存在的空白字符，并将内部空格替换为破折号
                # applicants = [applicant.rstrip().replace(" ", "-") for applicant in applicants]
                applicants = [re.sub(r"\s+", "-", applicant.strip())for applicant in applicants]
                # 格式化输出，连续字符之间用 '-' 连接起来
                applicants = '-'.join(applicants)

                utility_model_name_match = re.search(r"\(54\)使用外观设计的产品名称(.*?)(?=立体图|设计1|\Z)", text, re.S)
                
                info['专利权人'] = applicants if applicants else "Unknown"
                info['使用外观设计的产品名称'] = utility_model_name_match.group(1).strip() if utility_model_name_match else "Unknown"
            elif patent_type_match.group(1) == "实用新型专利":
                patent_type = "实用新型专利"
                # 提取实用新型专利相关信息
                # info['授权公告号'], info['授权公告日'] = re.findall(r'\bCN \d+ A\b', text)
                number_date_match = re.search(r'(CN \d+ [A-Z])\s*(\d{4}.\d{2}.\d{2})', text)
                if number_date_match:
                    info['授权公布号'] = number_date_match.group(1).replace(' ', '') 
                    info['授权公布日'] = number_date_match.group(2).replace('.', '-')  
                # applicants = re.findall(r"专利权人\s+((?:(?!地址).)*?)(?:地址|\s+\(\d{2}\)|$)", text)
                applicants = re.findall(r"专利权人\s+((?:(?!\(\d{2}\)).)*?)(?:地址|(?=\(\d{2}\))|$)", text)
                # 移除列表中每个字符串末尾可能存在的空白字符，并将内部空格替换为破折号
                # applicants = [applicant.rstrip().replace(" ", "-") for applicant in applicants]
                applicants = [re.sub(r"\s+", "-", applicant.strip())for applicant in applicants]
                # 格式化输出，连续字符之间用 '-' 连接起来
                applicants = '-'.join(applicants)

                utility_model_name_match = re.search(r"\(54\)实用新型名称(.*?)\(57\)", text, re.S)
                
                info['专利权人'] = applicants if applicants else "Unknown"
                info['实用新型名称'] = utility_model_name_match.group(1).strip() if utility_model_name_match else "Unknown"
        return patent_type, info
    def format_filename(self, type, info):
        # 使用提供的信息格式化新文件名
        if type == "发明专利申请":
            filename = f"{info['申请公布日']}_[{info['申请公布号'].replace(' ', '')}]-[{type}]-[{info['发明名称']}]-[{info['申请人']}]-[#patent]]"
            return filename
        elif type == "发明专利":
            filename = f"{info['授权公布日']}_[{info['授权公布号'].replace(' ', '')}]-[{type}]-[{info['发明名称']}]-[{info['专利权人']}]-[#patent]"
            return filename
        elif type == "外观设计专利":
            filename = f"{info['授权公布日']}_[{info['授权公布号'].replace(' ', '')}]-[{type}]-[{info['使用外观设计的产品名称']}]-[{info['专利权人']}]-[#patent]"
            return filename   
        elif type == "实用新型专利":
            filename = f"{info['授权公布日']}_[{info['授权公布号'].replace(' ', '')}]-[{type}]-[{info['实用新型名称']}]-[{info['专利权人']}]-[#patent]"
            return filename            

    def rename_pdf(self, pdf_path, new_name):
        directory, filename = os.path.split(pdf_path)
        new_filename = f"{new_name}.pdf"
        new_path = os.path.join(directory, new_filename)
        if os.path.exists(new_path) and (pdf_path != new_path):
            print(f"{new_filename} 当前文件已经存在目录中\n")
            self.move_pdf_to_duplicated_folder(pdf_path)
        elif os.path.exists(new_path) == False:
            os.rename(pdf_path, new_path)
            print('\n')

    def search_pdf_files(self, directory):
        pdf_files = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith(".pdf"):
                    pdf_files.append(os.path.join(root, file))
        return pdf_files

    def processPDFs(self):
        pdf_paths = self.search_pdf_files(self.folderPath)
        for pdf_path in pdf_paths:
            text = self.pdf_text_extractor(pdf_path)

            patent_type, patent_info = self.extract_patent_info(text)

            if patent_type == "Unknown":
                print("未知类型的专利，跳过处理")
                continue
            # 输出以确认
            # print("专利类型:", patent_type)
            # for key, value in patent_info.items():
            #     print(f"{key}: {value}")
            # print('\n')
            # 格式化文件名
            new_filename = self.format_filename(patent_type,patent_info)


            # # 打印预期的新文件名
            # print(new_filename)

            # 重命名专利
            self.rename_pdf(pdf_path,new_filename)
        
        print('所有PDF文件都已重命名。')

    # current_directory = folderPath
    # pdf_paths = search_pdf_files(current_directory)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())