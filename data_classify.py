import pandas as pd

#读数据
danpinData = pd.read_excel("附件2.xlsx")
pinleiData = pd.read_excel("附件1.xlsx")

print("读取成功")

df = danpinData.merge(pinleiData, on="单品编码", how="left")
df = df.sort_values(by=["分类编码", "单品编码", "销售日期", "扫码销售时间"])
print("排序完成")

#写入
with pd.ExcelWriter("result.xlsx", engine="openpyxl") as writer:
    for pinlei, pinleiGroup in df.groupby(["分类编码", "分类名称"]):
        output = []
        for danpin, danpinGroup in pinleiGroup.groupby("单品编码", sort=False):
            output.append(danpinGroup)
            output.append(pd.DataFrame([{}]))

        res = pd.concat(output, ignore_index=True)
        sheetName = str(pinlei[1])
        res.to_excel(writer, sheet_name=sheetName, index=False)
        print(f"已完成：{sheetName}")

print("任务完成")