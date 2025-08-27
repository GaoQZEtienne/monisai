import pandas as pd

#读数据
danpinData = pd.read_excel("附件2.xlsx")
pinleiData = pd.read_excel("附件1.xlsx")

print("读取成功")

df = danpinData.merge(pinleiData, on="单品编码", how="left")

df["销售日期"] = pd.to_datetime(df["销售日期"])
df["月份"] = df["销售日期"].dt.to_period("M")

summary = df.groupby(["分类编码", "分类名称", "单品编码", "单品名称", "月份"], as_index=False)["销量(千克)"].sum()

#按出现次数排序
summary["出现次数"] = summary.groupby("单品编码")["月份"].transform("count")
summary = summary.sort_values(by=["出现次数", "单品编码"], ascending=[False, True])

#写入
with pd.ExcelWriter("summary.xlsx", engine="openpyxl") as writer:
    for pinlei, pinleiGroup in summary.groupby(["分类编码", "分类名称"]):
        sheetName = str(pinlei[1])
        print(f"正在做：{sheetName}")
        pinleiGroup.to_excel(writer, sheet_name=sheetName, index=False)
print("任务完成")