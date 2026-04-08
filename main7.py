import pandas as pd

bank = pd.read_excel("银行流水新.xlsx",engine='openpyxl')
journal = pd.read_excel("银行日记账新.xlsx",engine='openpyxl')
    
amounts = []
amounts.extend(bank['收入'].tolist())
amounts.extend(bank['支出'].tolist())
amounts.extend(journal['借方金额'].tolist())
amounts.extend(journal['贷方金额'].tolist())

for amount in set(amounts):
    if pd.isna(amount) or amount == 0:
        continue
    
    b_i = (bank['收入'] == amount).sum()
    b_e = (bank['支出'] == amount).sum()
    j_d = (journal['借方金额'] == amount).sum()
    j_c = (journal['贷方金额'] == amount).sum()
    
    if b_i != j_c or b_e != j_d:
        print(f"金额{amount}: 流水(收{b_i},支{b_e})≠日记账(贷{j_c},借{j_d})")