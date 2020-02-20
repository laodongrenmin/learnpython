### xlrd
```
pip install xlrd
# Successfully installed xlrd-1.2.0

```
在表格中的文字还有格式
python – 如何在Excel文档单元格中找到文本子集的格式  
http://www.voidcn.com/article/p-amlcpxxy-bth.html

rb = xlrd.open_workbook(self.excel_file_path, formatting_info=True)
rs = rb.sheet_by_index(0)
