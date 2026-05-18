import pandas as pd
import numpy as np
pd.options.display.float_format="{:,.2f}".format
def test_data():
    """创建一个包含各种异常情况的测试数据集"""
    test_data={
        "金额": ["100.5", "-50", "200", "300.00", "12,000", "3,500,000,000", "1,000.50", "abc", "", None],
        "税率": ["0.13", "13%", "", None, "0.06", "1.5", "-0.1", "零点一", "0.1", 0.2],
        "交易日期": ["2024-01-01", "2024/01/02", "20240103", "", None, "Invalid Date", "Jan-15-2024", "2024-01-01", "2024-12-31", "2024-06-15"],
        "状态": ["processing", "paid", "cancelled", "pending", "未知状态", "", None, "处理中", "已支付", "已取消"],
        "所属部门": ["Sales", "Finance", "Tech", "", None, "人力资源部", "销售部", "财务部", "技术部", "运营部"],
        "联系电话": ["13800138000", "123-4567", "", None, "abcde", "123", "未提供", "13888888888", "13900139000", "13987654321"],
        "发票号码": ["INV001", "", None, "INV002", "INV-003", "INV004", "INV005", "INV006", "INV007", "INV008"],
        "备注": ["正常", "", "none", "null", "备注信息", None, "  ", "test", "备注", "无"]}
    return pd.DataFrame(test_data)
def load_data():
    source_file="进阶财务数据练习.xlsx"
    try:
        df=pd.read_excel(source_file)
        df.info()
        return df
    except FileNotFoundError:
        print(f"错误:未找到文件'{source_file}',请确认文件是否存在.")
        exit(1)
    except Exception as e:
        print(f"读取文件时发生错误:{e}")
        exit(1)
def diagnose_data(df):
    print("原始数据质量诊断报告")
    print("1.核心字段缺失值统计:")
    critical_cols=["金额","税率","交易日期","状态","所属部门"]
    null_counts=df[critical_cols].isnull().sum()
    for col in critical_cols:
        print(f"   - {col}: {null_counts[col]} 个空值")
    print("2.金额字段原始值抽样:")
    #查看非数值、负值等
    sample_non_numeric=df[~df["金额"].astype(str).str.match(r'^[\d.]+$',na=False)]["金额"].head()
    if not sample_non_numeric.empty:
        print(f"    发现非纯数字/小数点样本:{sample_non_numeric.tolist()}")
    else:
        print("    未发现非纯数字样本")
    print("3.状态字段原始值分布:")
    print(df["状态"].value_counts(dropna=False).head())
    print("4. 所属部门字段原始值分布:")
    print(df["所属部门"].value_counts(dropna=False).head())
    return df
def clean_amount(df):
    df["金额"]=df["金额"].astype(str).str.replace(r'[^\d.-]',"",regex=True)
    df["金额"]=pd.to_numeric(df["金额"],errors="coerce").round(2)
    df["金额_异常标记"]=df["金额"]>1e9  # 大于10亿标记为异常
    if df["金额_异常标记"].sum()>0:
        print(f"   警告: 发现 {df['金额_异常标记'].sum()} 条金额异常大的记录（>10亿）")
        print("   异常记录样本:")
        print(df[df["金额_异常标记"]].head())
    df["金额"]=df["金额"].fillna(0.0)
    return df
def calc_tax(df):
    df["税率"]=df["税率"].astype(str).str.replace(r'[^\d.]',"",regex=True)
    df["税率"]=pd.to_numeric(df["税率"],errors="coerce")
    df["税率"]=df["税率"].fillna(0.0)
    return df
def clean_phonenumber(df):
    df["联系电话"]=df["联系电话"].astype(str).str.replace(r'[^\d]',"",regex=True)
    df["联系电话"]=df["联系电话"].replace("",pd.NA)
    df["联系电话"]=df["联系电话"].fillna("未提供")
    return df
def clean_datetime(df):
    df["交易日期"]=pd.to_datetime(df["交易日期"],format="%Y-%m-%d",errors="coerce")
    print(df["交易日期"].dtype)
    failed_dates=df["交易日期"].isna().sum()
    if failed_dates>0:
        print(f"警告: {failed_dates} 条记录的日期解析失败")
        failed_samples = df[df["交易日期"].isna()][["交易日期"]].head()
        print("解析失败的日期样本:", failed_samples)
    return df
def clean_invoice(df):
    df["发票号码"]=df["发票号码"].astype(str).str.replace(r'[^\d]',"",regex=True)
    df["发票号码"]=df["发票号码"].replace("",pd.NA)
    df["发票号码"]=df["发票号码"].fillna("空")
    return df
def clean_remarks(df):
    df["备注"]=df["备注"].astype(str).str.strip().replace(r'^\s*$',pd.NA,regex=True)
    df["备注"]=df["备注"].replace(["nan","none","null","None"],pd.NA)
    df["备注"]=df["备注"].fillna("无")
    return df
def clean_status_and_department(df):
    df["状态"]=df["状态"].astype(str).replace({"processing":"处理中","paid":"已支付","cancelled":"已取消","pending":"未支付"})
    df["所属部门"]=df["所属部门"].astype(str).replace({"Sales":"销售部","Finance":"财务部","Tech":"技术部","HR":"人事部","Marketing":"市场部"})
    return df
def analysis(df):
    dept_cost=df.groupby("所属部门")["金额"].sum().round(2)
    print(dept_cost)
    df["月份"]=df["交易日期"].dt.to_period("M").astype(str)
    monthly_sales=df.groupby("月份")["金额"].sum().round(2)
    print(monthly_sales)
    status_summary=df.groupby("状态")["金额"].agg(订单金额合计="sum",订单数量="count").round(2)
    print(status_summary)
    return df
def self_check(df):
    assert (df["金额"]>=0).all(),f"金额出现负值,共{(df['金额']<0).sum()} 条"
    assert df["状态"].isin(["处理中","已支付","已取消","未支付"]).all(),"状态异常"
    assert df["所属部门"].notna().all(),f"所属部门有{df['所属部门'].isna().sum()}个空值"
    assert df["金额_异常标记"].sum()==0,f"金额存在异常,共 {df['金额_异常标记'].sum()} 条记录"
    print("自测通过:数据质量符合预期")
    return df
def check(df):
    print(df.isnull().sum())
    df.to_excel("进阶财务数据练习清洗后数据.xlsx",index=False)
    print("财务数据清洗与分析完成，结果已导出")
if __name__ == "__main__":
    run_mode=input("运行模式？(1: 使用测试数据, 2: 使用真实文件): ")
    if run_mode == "1":
        print("进入测试模式...")
        df=test_data()
    else:
        print("进入真实数据处理模式...")
        df=load_data()
    df=diagnose_data(df)
    df=clean_amount(df)
    df=calc_tax(df)
    df=clean_phonenumber(df)
    df=clean_datetime(df)
    df=clean_invoice(df)
    df=clean_remarks(df)
    df=clean_status_and_department(df)
    df=analysis(df)
    df=self_check(df)
    check(df)

