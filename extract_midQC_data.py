import pandas as pd
import numpy as np
from collections import Counter
from docx import Document
from docx.shared import Cm, Mm

class MidQcExtracter:
    def __init__(self, file=None, plate=None):
        self.file = file
        self.plate = plate
        self.data = None

    def __del__(self):
        pass

    def update_file_and_plate(self, file, plate):
        self.file = file
        self.plate = plate
        self.data = None

    def get_df(self):
        df = pd.read_excel(io=self.file)

        self.data = df[df['孔板号'] == self.plate]
        # self.data = df[df['孔板号'] == self.plate].reset_index()

        self.data = self.data.copy()

        self.data['浓度'] = round(self.data['浓度'], 2)

        return self.data

    def get_size(self):
        return len(self.data)

    def get_risk_list(self):
        risk_list = []

        col_lst = self.data['结果'].tolist()
        for ind, val in enumerate(col_lst):
            if val == '风险上机':
                risk_list.append(ind+1)
        return risk_list

    def get_risk_counter(self):
        risk_dic = {}

        col_lst = self.data['结果'].tolist()
        while np.nan in col_lst:
            col_lst.remove(np.nan)
        risk_dic = Counter(col_lst)

        return risk_dic


if __name__ == '__main__':
    file_name = 'C:/Users/ben/PycharmProjects/中间质控汇总 2021.3.25(1).xls'

    try:
        mid_qc = MidQcExtracter()
        mid_qc.update_file_and_plate(file_name, 'TZ0028')
        data = mid_qc.get_df()
        mid_risk_dic = mid_qc.get_risk_counter()
        print(data)

        print(mid_risk_dic)
    except Exception as e:
        print(str(e))

    if 1:
        doc = Document()  # 新建文档对象
        # 增加表格
        colnames_1 = ["孔板号", "孔号", "样本编号", "稀释后浓度（吸光度法）ng/μl", "结果"]
        colwidth_1 = [20, 25, 35, 50, 35]
        # tablestyle在模板里指定
        sample_table = doc.add_table(rows=1, cols=0, style='Table Grid')
        sample_table.autofit = False
        # 设置列宽
        for i in range(len(colnames_1)):
            sample_table.add_column(width=Mm(colwidth_1[i]))

        # 设置表头
        hdr_cells = sample_table.rows[0].cells
        for i in range(len(colnames_1)):
            hdr_cells[i].text = colnames_1[i]

        # 填充数据
        for index, row in data.iterrows():
            row_cells = sample_table.add_row().cells  # add new row
            i = 0
            for item in row:
                item_str = str(item)
                if item_str == "nan":
                    item_str = ''
                row_cells[i].text = item_str
                i += 1
        # 保存文件
        doc.save(f'CAS-CN1基因芯片实验报告-testing.docx')
