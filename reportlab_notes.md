### 一、 Reportlab 学习记录
笑话： 高高兴兴安装了库，就是不能使用，怎么检查都安装好了的，就是找不到库，百度、google找不到解决方法。
```
这样安装的 
pip install reportlib  
这样使用的
from reportlab.pdfgen import canvas
使劲看，是安装错误了包，那个字母不一样，功能差别大了去了
```

1、 中文字体问题
- 问题出现
```python
# 一个HelloWorld很容易成功，心情很好，工具上手容易
from reportlab.pdfgen import canvas
def hello(c):
    c.drawString(100,100,"Hello World")
c = canvas.Canvas("hello.pdf")
hello(c)
c.showPage()
c.save()

# 但是当输入，你好，整个人心里就凉了

c.drawString(100,100,"Hello World")
# 改为
c.drawString(100,100,"你好！")
# 三个黑框框出现了
```
- 解决方法

pyhton之Reportlab模块 https://www.cnblogs.com/hujq1029/p/7767980.html ，描述的比较清楚，是因为没有注册中文字体。

python+reportlab学习：解决中文问题 https://blog.csdn.net/jtscript/article/details/45027997

```python
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('仿宋', 'simfang.ttf'))
def hello(c):
    c.c.setFont("仿宋", 40)
    c.drawString(100,100,"您好！")

```
可以完美输出中文了， 网上还有人建议复制字体文件到工程下，其实是不需要，TTFont会自动查找系统的字体文件库，
这里有个问题就是。 比如要把excel文件导出生成PDF，字体名称是中文的仿宋。但是这里的这个仿宋，是我手工输入的，
reportlab 导入的是英文的名称，在windows的字体目录下，看到的中文名称就是仿宋，怎么得到这个中文名称呢。

- 曲线得到中文名称的方法

获取TrueType字体信息 https://blog.csdn.net/kwfly/article/details/50986338 对TrueType字体文件说的比较清楚，
通过调试程序，看到，TTFont类的实现并没有取当前系统页对应的名称，而是固定取得英文的名称， 一种方法是修改
TTFont类，解析出中文名称，这样破坏了原来的文件，以后升级有问题。 二种是按照上面的方法去实现，找到，感觉
比较麻烦，还有三种就是建立一个对照表，自己维护，穷尽字体文件。选择方法三简单实现耶罗。