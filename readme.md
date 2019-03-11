# 大批量的从pdf文件导出表格数据

## 项目简介
该项目是为解决大批量的从pdf文件导出表格，采取的方案是[tabula](https://github.com/tabulapdf/tabula)+自定义调整。本项目代码就是自定义调整的代码实现。
## 原理简述
对不规则的复杂表格输出时，tabula会出现部分内容错乱的情况。经过对tabula的json导出文件进行分析，该文件保留了源PDF文件每行每列的位置信息，分别是用top,height,left,width表示，并且同行的height和top值相等，同列的width+left = 对应标题列的width+left。这是本项目的实现原理。
## 实现过程
指定表格的标题行，根据行列规则确定行列边界，逐一扫描每一个单元格计算并标记该单元格正确的行列坐标。然后对同样行列的值，根据目标文件的语义判断是先进行行方向内容合并还是列方向内容合并。

## 安装组件：
1. [安装Java](http://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html "java")运行环境。目的是能够运行tabula的jar包。
2. [安装Python3](https://www.python.org/downloads/ "下载Python3")
3. [安装PyPDF2](https://pypi.org/project/PyPDF2/ "PyPDF2")
4. [安装xlwt](https://pypi.org/project/xlwt/ "xlwt")

## 必要的修改：
1. 添加标题行。请在代码头部找到变量table_header，往其内容中添加你要调整的文档的标题行。格式参照现有数据。这里的标题行不是通常表格的第一行，准确的说是为表格列提供调整列基准的行。参照target/baogaobiao.pdf文件的标题行应该是：['名称代码','坐标','航路']，而不是：['报告表']。

## 使用方式：
```python
python adjust_table.py <pdf_file>|<directory> [check=yes(default)|no] [page=2|2-10|all|top5(default)] [mode=lattice(default)|stream]
```
**check**:表示对pdf文件中每一页的表格单独输出一个json和excel文件，方便检查对照转换效果。默认是yes，输出目录是./output。

**page**:指定转换页。
例如：
* python adjust_table.py target/G.pdf 
* python adjust_table.py target/  check=no 
* python adjust_table.py target/baogaobiao.pdf  check=yes page=5-10
* python adjust_table.py target/ check=no mode=lattice

## 有待改进
转换效果因文档不同还有待改善，并且不同的文档可能有需要特别自定义的地方。暂时整理到这里，有时间再更新。
