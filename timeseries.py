import pandas as pd

#读数据
all_sheets = pd.read_excel("summary.xlsx", sheet_name=None) 
DF = []
for sheet, df in all_sheets.items():
    if sheet != "品类统计":
        DF.append(df)
print("读取成功")

summary = pd.concat(DF, ignore_index=True)

summary["月份"] = pd.to_datetime(summary["月份"]).dt.to_period("M")
time_index = pd.period_range("2020-07", "2023-06", freq="M")
pivot = summary.pivot_table(
    index="月份",
    columns=["分类名称", "单品名称"],
    values="销量(千克)",
    aggfunc="sum",
    sort=False
)
pivot = pivot.reindex(time_index)

counts = (
    summary.groupby(["分类名称", "单品名称"], as_index=False)["出现次数"]
           .max()
           .sort_values(["分类名称", "出现次数", "单品名称"], ascending=[True, False, True])
)

ordered_cols = [(r["分类名称"], r["单品名称"]) for _, r in counts.iterrows() if r["出现次数"] >= 30]

pivot = pivot.reindex(columns=pd.MultiIndex.from_tuples(ordered_cols))

#写入
pivot.to_excel("TimeSeries.xlsx")
