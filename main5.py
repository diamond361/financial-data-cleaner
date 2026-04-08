import pandas as pd

# 最简单的一键核对
bank = pd.read_excel("银行流水新.xlsx",engine='openpyxl')
journal = pd.read_excel("银行日记账新.xlsx",engine='openpyxl')

# 计算差额
income_diff = bank['收入'].sum() - journal['贷方金额'].sum()
expense_diff = bank['支出'].sum() - journal['借方金额'].sum()

print(f"收入差异: {income_diff:,.2f} 元")
print(f"支出差异: {expense_diff:,.2f} 元")

if abs(income_diff) < 0.01 and abs(expense_diff) < 0.01:
    print("✅ 金额核对通过！")
else:
    print("❌ 金额不匹配！")