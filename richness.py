import pandas as pd

#读数据
df = pd.read_excel("附件2.xlsx")

df["销售日期"] = pd.to_datetime(df["销售日期"])
df["月份"] = df["销售日期"].dt.to_period("M")

richness = df.groupby("月份")["单品编码"].nunique().reset_index()
richness.rename(columns={"单品编码": "菜品丰富度"}, inplace=True)
#写入
richness.to_excel("richness.xlsx",  sheet_name="Sheet1", index=False)