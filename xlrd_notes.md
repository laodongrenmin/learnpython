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


xlrd简介  https://blog.csdn.net/weixin_32759777/article/details/84998479

python使用xlrd读取excel数据时，整数变小数的解决办法 https://www.cnblogs.com/wanglei-xiaoshitou1/p/9401261.html

Excel 日期时间格式讲解 https://blog.csdn.net/lipinganq/article/details/78436089