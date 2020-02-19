### 一、 Office2PDF
在 windows 系统（需安装office套件）中，利用 python 的 win32com 包
实现对Office文件的操作，转换为pdf文件。
支持 doc, docx, ppt, pptx, xls, xlsx 等格式
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

### 二、 PDF2Img
利用纯Python库，转换PDF到图像，选择了一个比较高效的库， 效果还不错，速度非常快
#### 1 环境搭建
```
pip install PyMuPDF
# Successfully installed PyMuPDF-1.16.10
# 安装的时候指定 镜像源https://pypi.tuna.tsinghua.edu.cn/simple，否则会很慢
```
#### 2 程序代码
```buildoutcfg
见 src/pdf2img
```
#### 3 效果完美，转换效率高
```buildoutcfg
cost time: 193.21322441101074 ms
# 两页1056x816的图像，耗时193毫秒，非常理想
```



