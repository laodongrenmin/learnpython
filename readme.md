### Office2PDF
在 windows 系统（需安装office套件）中，利用 python 的 win32com 包可以实现对Office文件的操作，可以批量转换为pdf文件。支持 doc, docx, ppt, pptx, xls, xlsx 等格式
#### 1 环境搭建
```
conda create -n learn python=3.7 
pip install pywin32
# 安装的时候指定 镜像源https://pypi.tuna.tsinghua.edu.cn/simple，否则会很慢
```
#### 2 程序代码
```buildoutcfg
见 src/Office2PDF
```

