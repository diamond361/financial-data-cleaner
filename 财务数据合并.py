import pandas as pd
import numpy as np
import visualization as vs
def merged_data():
    try:
        orders_df=pd.read_excel("财务数据合并练习.xlsx",sheet_name="订单表")
        customers_df=pd.read_excel("财务数据合并练习.xlsx",sheet_name="客户表")
        products_df=pd.read_excel("财务数据合并练习.xlsx",sheet_name="产品表")
        departments_df=pd.read_excel("财务数据合并练习.xlsx",sheet_name="部门表")
    except FileNotFoundError:
        print('错误:未找到文件"财务数据合并练习.xlsx",请确认文件是否存在')
        return None
    except Exception as e:
        print(f"读取文件时发生错误:{e}")
        return None
#左连接订单表和客户表
    try:
        merged_step1_df=pd.merge(orders_df,customers_df,on="客户ID",how="left")
        merged_step2_df=pd.merge(merged_step1_df,products_df,on="产品ID",how="left")
        customer_level_summary = merged_step2_df.groupby("客户等级")["金额"].sum()
        customer_level_df = customer_level_summary.reset_index()
    except Exception as e:
        print(f"数据处理失败:{e}")
        return None
    try:
        with pd.ExcelWriter("合并后数据.xlsx",engine="openpyxl") as writer:
            merged_step2_df.to_excel(writer,sheet_name="合并后数据",index=False)	
            customer_level_df.to_excel(writer,sheet_name="客户等级汇总",index=False)
        print("所有分析结果已保存到:合并后数据.xlsx")
        return merged_step2_df
    except PermissionError:
        print("保存失败:文件被占用或无写入权限,请关闭文件后重试")
        return None
    except Exception as e:
        print(f"保存失败:{e}")
        return None

def analyze_and_visualize(df):
    if df is None or df.empty:
        print("错误，没有有效数据可供分析")
        return False 
    print("开始数据分析和可视化")
    if "成本价" in df.columns:
        df["利润"] = df["金额"] - df["成本价"]
        df["利润率"] = (df["利润"]/df["金额"]*100).round(2)
    print("正在生成可视化图表")
#生成各地区销售额分布图
    if "所在地区" in df.columns and "金额" in df.columns:
        vs.plot_sales_by_region(df)
# 生成产品类别销售占比图
    if "产品类别" in df.columns and "金额" in df.columns:
        vs.plot_category_pie(df)
    print("数据分析和可视化完成！")
    return True

if __name__ == "__main__":
    merged_df = merged_data()
    if merged_df is not None:
        analyze_and_visualize(merged_df)
        print(f"总销售额: {merged_df['金额'].sum():,} 元")
        print(f"总订单数: {len(merged_df)} 笔")
        if '利润' in merged_df.columns:
            print(f"总利润: {merged_df['利润'].sum():,} 元")
            print(f"平均利润率: {merged_df['利润率'].mean():.2f}%")
        if '客户等级' in merged_df.columns:
            level_stats = merged_df.groupby('客户等级')['金额'].agg(['count', 'sum', 'mean'])
            level_stats.columns = ['订单数', '总金额', '平均金额']
            print(level_stats)
        if '产品类别' in merged_df.columns:
            category_stats = merged_df.groupby('产品类别')['金额'].agg(['count', 'sum', 'mean'])
            category_stats.columns = ['订单数', '总金额', '平均金额']
            print(category_stats)
    else:
        print("数据合并失败，程序退出")
