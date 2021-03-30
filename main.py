import numpy as np
import sys
import models
import excel


def main():
    sp_size = 10    # 样本的元素个数
    sp_cap = 200    # 样本容量
    filename = sys.argv[1]
    data_sheet = excel.DataSheet(filename)
    data_sheet.entitle()
    data_sheet.format()

    for row in range(3, sp_cap + 3):
        sample = np.random.randint(101, 1000, sp_size)
        estimate = ( 
                     models.average_1(sample),
                     models.median_2(sample),
                     models.symmetry_3(sample),
                     models.avg_interval_4(sample, sp_size),
                     models.evenly_split_intvl_5(sample, sp_size)
                   )
        data_sheet.write(row, sample, sp_size, estimate)



if __name__ == '__main__':
    main()

