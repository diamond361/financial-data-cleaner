import pandas as pd

ledger = pd.read_excel('银行日记账新.xlsx')
statement = pd.read_excel('银行流水新.xlsx')

ledger['实际金额'] = ledger['贷方金额'] - ledger['借方金额']
statement['实际金额'] = statement['收入'] - statement['支出']

matches = []
for i, ledger_row in ledger.iterrows():
    ledger_amount = ledger_row['实际金额']
    ledger_company = str(ledger_row['对方单位'])
    
    for j, stmt_row in statement.iterrows():
        stmt_amount = stmt_row['实际金额']
        stmt_company = str(stmt_row['对方户名'])
        
        # 金额相同
        if abs(ledger_amount - stmt_amount) < 0.01:
            # 检查公司名称是否相似
            if (ledger_company in stmt_company) or (stmt_company in ledger_company):
                status = '✅ 金额+公司都匹配'
            else:
                status = '⚠️ 金额相同但公司不同'
            
            matches.append({
                '状态': status,
                '日记账公司': ledger_company,
                '日记账金额': ledger_amount,
                '流水公司': stmt_company,
                '流水金额': stmt_amount
            })

# 统计结果
print("匹配结果统计:")
for status in ['✅ 金额+公司都匹配', '⚠️ 金额相同但公司不同']:
    count = len([m for m in matches if m['状态'] == status])
    print(f"{status}: {count}笔")

pd.DataFrame(matches).to_excel('匹配检查.xlsx', index=False)