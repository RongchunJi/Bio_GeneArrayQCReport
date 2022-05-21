import pandas as pd

table_lst = [['起始蓝色冻存管', ' ', ' ', '客户样本板'],
             ['起始蓝色冻存管', 'PCR板', ' ', 'DNA样本均一化稀释处理'],
             ['PCR板', '深孔板', ' ', '芯片扩增片段化干燥重悬'],
             ['深孔板', 'PCR板', ' ', '重悬后转板加杂交液后续做质控及变性处理'],
             ['PCR板', '酶标板', ' ', '重悬液质控前稀释及浓度测定'],
             ['酶标板', '点样板', ' ', '中间过程电泳质控'],
             ['PCR板', '杂交板（带条形码扫码上机）', ' ', '转至杂交板上Titan杂交']]


class PlatesExtracter:
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
        df = pd.read_csv(self.file, header=None)
        # print(df)

        lst = df[0].tolist()

        for index, item in enumerate(lst):
            table_lst[index][2] = str(item)

        return pd.DataFrame(table_lst)


if __name__ == '__main__':
    file_name = 'C:/Users/ben/PycharmProjects/TZ0028/plates_name.txt'
    plate_name = 'TZ0028'

    plates = PlatesExtracter()
    plates.update_file_and_plate(file_name, plate_name)

    ls = plates.get_lst()

    print(ls)
