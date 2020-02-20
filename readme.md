### 一、 Office2PDF
在 windows 系统（需安装office套件）中，利用 python 的 win32com 包
实现对Office文件的操作，转换为pdf文件。
支持 doc, docx, ppt, pptx, xls, xlsx 等格式
#### 1 环境搭建
```
conda create -n learn python=3.7 
pip install pywin32
# 安装的时候指定 镜像源https://pypi.tuna.tsinghua.edu.cn/simple，否则会很慢
# office 用的是office 2019版， 红米笔记本装机赠送的学生版
```
#### 2 程序代码
```
# 详见 src/Office2PDF
# excel to pdf 代码如下
from win32com.client import DispatchEx
xlApp = DispatchEx("Excel.Application")
books = xlApp.Workbooks.Open('excel.xls', False)
books.ExportAsFixedFormat(0, 'excel.pdf')
```

#### 3 效果非常完美，就是转换效率蛮低
```
    # 时间单位是秒
    open excel app cost: 2.306511163711548
    convert excel to pdf one time cost: 6.887148380279541
    convert excel to pdf one time cost: 7.744565725326538
    quit excel app cost: 20.519250631332397
    open app and convert and quit app , excel to pdf one time cost: 11.418348789215088
    open word app cost: 5.517384767532349
    convert word to pdf one time cost: 4.319069862365723
    open ppt app cost: 3.7118749618530273
    convert ppt to pdf one time cost: 4.690080642700195
```
#### 4 遇到的一些问题记录
4.1、 excel 转 PDF， 设置一个sheet页转换为一页PDF， 网上找半天没发现。后来通过到execl里面录制宏找到参数，解决
```python
from win32com.client import DispatchEx
xlApp = DispatchEx("Excel.Application")
books = xlApp.Workbooks.Open('excel.xls', False)
sheet =  sheet = books.Sheets(1)   #  这里序号是从1开始的
# 设置为 将工作表调整为一页， 默认是 无缩放
sheet.PageSetup.Zoom = False
sheet.PageSetup.FitToPagesWide = 1
sheet.PageSetup.FitToPagesTall = 1
# 更多的参数设置查询微软文档，非常详尽 https://docs.microsoft.com/zh-cn/office/vba/api/excel.workbook.exportasfixedformat
# sheet.ExportAsFixedFormat(0, 'excel.pdf')  导出一个sheet
books.ExportAsFixedFormat(0, 'excel.pdf')    # 导出全部sheet  
```
4.2、 全部sheet表都设置 调整为一页， 现在都没有找到好方法，网上有例子是 用xlrd读excel然后获取count，感觉不太好。 
我用了try , 异常以后就不设置了。

4.3、 ppt 转 PDF，ppt原始文件里面的页，设置为了隐藏，导不出来，提示异常信息很清楚，由于是学习，没有好注意，网上到处找原因，没有。
估计是提示太清楚了，没有人会不注意的。学习的时候，一定要注意看异常信息。

4.4、 ppt 转 PDF， 报告 -2147024894 错误，就这一个信息，不知道原因，上网查询很久才找到，这个意思是 //系统找不到指定的文件。
原来是我的输入文件路径写错了。
### 二、 PDF2Img
利用纯Python库，转换PDF到图像，选择了一个比较高效的库， 效果还不错，速度非常快
#### 1 环境搭建
```
pip install PyMuPDF
# Successfully installed PyMuPDF-1.16.10
# 安装的时候指定 镜像源https://pypi.tuna.tsinghua.edu.cn/simple，否则会很慢
```
#### 2 程序代码
```
# 见 src/pdf2img
import fitz
pdf_doc = fitz.open(pdf_file_path)
# 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
# 此处若是不做设置，默认图片大小为：792X612, dpi=96
mat = fitz.Matrix(1.3, 1.3).preRotate(int(0))  
pdf_doc[0].getPixmap(matrix=mat, alpha=False).writePNG('test.png')  # 第一页
```
#### 3 效果完美，转换效率高
```
cost time: 166.52226448059082 ms
# 两页1029x795的图像，耗时 166 毫秒，非常理想
```

### 三、 ReportLab
一个生成PDF的python类库
#### 1 环境搭建
```
pip install reportlab
# Successfully installed pillow-7.0.0 reportlab-3.5.34
# 安装的时候指定 镜像源https://pypi.tuna.tsinghua.edu.cn/simple，否则会很慢
```
#### 2 程序代码
```
from reportlab.pdfgen import canvas
def hello(c):
    c.drawString(100,100,"Hello World")
c = canvas.Canvas("hello.pdf")
hello(c)
c.showPage()
c.save()
```


