import pandas as pd
import numpy as np
source_file="财务数据清洗练习.xlsx"
df=pd.read_excel(source_file,sheet_name="Sheet2")
df.info()  #数据检查
df["客户姓名"] = df["客户姓名"].replace(r'^\s*$', np.nan, regex=True)
df["销售员"] = df["销售员"].replace(r'^\s*$', np.nan, regex=True)
df["客户姓名"]=df["客户姓名"].fillna("未知客户")
df["销售员"]=df["销售员"].fillna("未分配")
df.dropna(axis=0,how="any",subset=["订单编号"],inplace=True)
print("订单编号缺失:",df["订单编号"].isnull().sum())
print("客户姓名缺失:",df["客户姓名"].isnull().sum())
print("销售员缺失:",df["销售员"].isnull().sum())
print("客户姓名空字符串:", (df["客户姓名"].astype(str).str.strip() == '').sum())
print("销售员空字符串:", (df["销售员"].astype(str).str.strip() == '').sum())
df.to_excel("清洗后数据.xlsx",index=False)
