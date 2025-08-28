import pandas as pd
from scipy.stats import jarque_bera


df = pd.read_excel("TimeSeries.xlsx",skiprows=1)

res = {}
for col in df.columns:
    data = df[col].dropna()
    if data.dtype.kind in 'biufc': 
        jb_stat, p_value = jarque_bera(data)
        res[col] = {"JB统计量": jb_stat, "p值": p_value}

results = pd.DataFrame(res).T

results.to_excel("JB检验.xlsx")
print("任务完成")