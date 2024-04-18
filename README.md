> 项目需要，好像还没人做过，没有参考资料
## 任务介绍
![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/2444355a32514eaab64fee08b44945ca.png)
![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/31218326954a44cfb7934ee132f489de.png)

把非结构化数据变为结构化数据，无线电频率划分规定[下载链接](https://wap.miit.gov.cn/gyhxxhb/jgsj/cyzcyfgs/bmgz/wxdl/art/2023/art_1e98823e689f42ca9ed14dcb6feec07a.html)
任务是把word文档转存在数据库中方便管理。
工具：python+docx+pandas
需要处理一下原始数据：
因为khz/Mhz/Ghz都在一起了
处理后的目录：
![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/243899e05c4e48558504a565a63ce1bb.png)
代码和处理后的文件放在我的[github仓库](https://github.com/youthfulever/read_frequency_band)
直接代码：

```bash
from docx import Document
import pandas as pd

list_ = []  # 初始化一个空列表，用来装后面的数据字典


# 处理khz
path = "khz.docx"
docx = Document(path)
for table in docx.tables:  # 循环所有的表格
    row_counter = 0  # 初始化行计数器
    for row in table.rows:  # 循环表格中的所有行
        row_counter += 1  # 每次循环时，增加行计数器
        if row_counter <= 2:  # 如果是前两行，则跳过
            continue
        cells = row.cells  # 获取当前行的所有单元格
        # 假设每个单元格中的内容都是通过换行符分隔的
        # 并且第一列包含了起始频率和结束频率，用"—"分隔
        # 其余列包含了业务信息
        if cells:
            item = cells[0].text
            dict_ = {
                '起始频率(KHz)': str(item.split("\n")[0]).split("—")[0],
                '结束频率(KHz)': str(item.split("\n")[0]).split("—")[1],
                '业务': ','.join(item.split("\n")[1:])
            }
            list_.append(dict_)  # 将字典添加到列表中


# Mhz,需要成1000
# 处理khz
path = "mhz.docx"
docx = Document(path)
for table in docx.tables:  # 循环所有的表格
    row_counter = 0  # 初始化行计数器
    for row in table.rows:  # 循环表格中的所有行
        row_counter += 1  # 每次循环时，增加行计数器
        if row_counter <= 2:  # 如果是前两行，则跳过
            continue
        cells = row.cells  # 获取当前行的所有单元格
        # 假设每个单元格中的内容都是通过换行符分隔的
        # 并且第一列包含了起始频率和结束频率，用"—"分隔
        # 其余列包含了业务信息
        if cells:
            item = cells[0].text
            temp=str(item.split("\n")[0]).replace(' ', '')
            dict_ = {
                '起始频率(KHz)': float(temp.split("—")[0])*1000,
                '结束频率(KHz)': float(temp.split("—")[1])*1000,
                '业务': ','.join(item.split("\n")[1:])
            }
            list_.append(dict_)  # 将字典添加到列表中


# GHZ 需要乘以1000 *1000
path = "Ghz.docx"
docx = Document(path)
for table in docx.tables:  # 循环所有的表格
    row_counter = 0  # 初始化行计数器
    for row in table.rows:  # 循环表格中的所有行
        row_counter += 1  # 每次循环时，增加行计数器
        if row_counter <= 2:  # 如果是前两行，则跳过
            continue
        cells = row.cells  # 获取当前行的所有单元格
        # 假设每个单元格中的内容都是通过换行符分隔的
        # 并且第一列包含了起始频率和结束频率，用"—"分隔
        # 其余列包含了业务信息
        if cells:
            item = cells[0].text
            temp=str(item.split("\n")[0]).replace(' ', '')
            dict_ = {
                '起始频率(KHz)': float(temp.split("—")[0])*1000*1000,
                '结束频率(KHz)': float(temp.split("—")[1])*1000*1000,
                '业务': ','.join(item.split("\n")[1:])
            }
            list_.append(dict_)  # 将字典添加到列表中

# 将列表转换为DataFrame
df = pd.DataFrame(list_)

# 将DataFrame保存为Excel文件
excel_path = "alldata.xlsx"
df.to_excel(excel_path, index=False)

print(f"数据已保存至 {excel_path}")
```
得到excel数据如下：
![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/9bcba68c69aa442d88ff399bc9e52e5b.png)
参考[地址](https://blog.csdn.net/D_xiaoniu/article/details/106743304)读入mysql数据库，展示如下：
![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/091028569c82421dbfd09d651024efed.png)

