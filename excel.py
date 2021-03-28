import openpyxl


class DataSheet():
    """建模数据表格, 有创建和读写功能"""

    def __init__(self):
        self.rows = 201
        self.columns = 16
        self.filename = 'model-data.xlsx'
        self.wb = self._init_wb(self.filename)
        self.sheet = self.wb['Sheet']


    def _init_wb(self, filename):
        """初始化工作簿"""
        try: 
            wb = openpyxl.load_workbook(filename)
        except FileNotFoundError:
            wb = openpyxl.Workbook()

        return wb


    def write(self, row, sample, estimate):
        """把样本与估测结果写入指定行"""
        for i in range(self.columns - 1):
            locat = f'{openpyxl.utils.get_column_letter(i+2)}{row}'
            if 0 <= i <= 9:
                self.sheet[locat] = sample[i]
            elif 10 <= i <= (self.columns - 1):
                self.sheet[locat] = estimate[i - 10]
            else:
                raise IndexError(f'No Such i = {i}')
        self.wb.save(self.filename)
