import numpy as np


def average_1(sample):
    """模型一: 平均值模型"""
    return int(2 * np.mean(sample) - 1)


def median_2(sample):
    """模型二: 中位数模型"""
    return int(2 * np.median(sample) - 1)


def symmetry_3(sample):
    """模型三: 两端间隔对称模型"""
    return (sample.max() + sample.min() - 1)


def avg_interval_4(sample, sp_size):
    """模型四: 平均间隔模型"""
    return int((1 + 1/sp_size) * sample.max() - 1)


def evenly_split_intvl_5(sample, sp_size):
    """模型五: 区间均分模型"""
    return int((1 + 1/(2 * sp_size - 1)) * (sample.max() - 1/(2 * sp_size)))
