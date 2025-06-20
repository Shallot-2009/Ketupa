import pandas as pd
from sklearn.model_selection import train_test_split

def load_data(file_path):
    """加载数据集"""
    data = pd.read_csv(file_path)
    return data

def split_data(data):
    """划分训练集和测试集"""
    train_data, test_data = train_test_split(data, test_size=0.2, random_state=42)
    return train_data, test_data