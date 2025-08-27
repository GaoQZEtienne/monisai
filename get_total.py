import pandas as pd

#读数据
danpinData = pd.read_excel("附件2.xlsx")
pinleiData = pd.read_excel("附件1.xlsx")

print("读取成功")

df = danpinData.merge(pinleiData, on="单品编码", how="left")

df["销售日期"] = pd.to_datetime(df["销售日期"])
df["月份"] = df["销售日期"].dt.to_period("M")

summary = df.groupby(["分类编码", "分类名称","月份"], as_index=False)["销量(千克)"].sum()

with pd.ExcelWriter("summary.xlsx", mode="a", engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="品类统计", index=False)