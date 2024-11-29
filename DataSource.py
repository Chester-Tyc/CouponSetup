import pandas as pd


def read_excel(source):
    # 一次性读取整个表格
    df = pd.read_excel(source)
    print(df)

    return df
