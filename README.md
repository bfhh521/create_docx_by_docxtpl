# create_docx_by_docxtpl
# 概述

在工频电磁场测量仪校准实验中，需要根据委托信息、实验数据来编制委托协议、客户沟通记录，以及出具实验报告。
大概流程是从现有的excel表格中抽取一系列字符，然后填写到word固定位置后批量生成文档，手动操作繁琐，重复劳动缺乏效率，使用docxtpl库实现。

 **应用环境如下：**
* Windows 10
* Python 3.6
* docxtpl-0.6.3

# 安装支持
```python
conda install docxtpl
#或
pip install docxtpl
#再不行
conda install -c conda-forge docxtpl
#docxtpl安装，会自动安装依赖库docx和jinja2
```
docxtpl：[官方文档](https://docxtpl.readthedocs.io/en/latest/) 　 [github](https://github.com/elapouya/python-docx-template)
jinja2：[官方文档](https://jinja.palletsprojects.com/en/2.10.x/)　 [中文文档](http://docs.jinkan.org/docs/jinja2/templates.html#import-visibility)

# 实现思路
1.jinja2使用{{...}}声明模版中的变量，我们将docx模版中需要替换的内容使用{{...}}手动标注起来。
2.从xls中读取需要替换的值，并与docx模版中预设的变量名对应起来。
3.使用docxtpl库中的DocxTemplate.render完成模板替换。
4.输出替换后的docx。

# 准备模版
将需要替换的位置使用双大括号进行标准，并添加变量名。
这里需要注意的是这里应该对{{var}}本身的文本格式完成调整，这样后面替换时就不需要再单独对文本格式进行处理了。
需要单独调整的可以通过docxtpl库使用富文本的方式操作。
![docx模版](https://img-blog.csdnimg.cn/20200311231520577.png)
# 读取委托信息
中间发生过一个问题，因为手机号信息输出时需要整形，使用pd.astype('int64')时,发现对存在“NAN”的数据无法处理，网上查到说[pandas 0.24以上的版本已经可以支持了](https://stackoverflow.com/questions/11548005/numpy-or-pandas-keeping-array-type-as-integer-while-having-a-nan-value)，就去升级了pandas到1.0.1，结果spyder打不开了。
![spyder打不开了](https://img-blog.csdnimg.cn/20200312184819578.png)
又查到好像说是升级pandas时，依赖库把mkl升级到了2018.0.3，而这个版本有问题，[(参考链接)](https://github.com/spyder-ide/spyder/issues/7357)，建议重新装回mkl 2018.0.2。
```python
conda install mkl=2018.0.2
```
装完后，出现第二个问题。
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200312221228827.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2JmaGg1MjE=,size_16,color_FFFFFF,t_70)
[问题解决：
因为提示少了”IPython.core.inputtransformer2“模块，所以找到对应的文件夹
”D:\anaconda3\Lib\site-packages\IPython\core“
发现在这下面的文件与可以正常运行的ipython文件夹对比少了”inputtransformer2.py“和”async_helpers.py“两个文件，从中复制过来,正常打开即可~~~](https://blog.csdn.net/Y_yuxiaoyu/article/details/103792381)

然后是第三个问题，
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200312221530823.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2JmaGg1MjE=,size_16,color_FFFFFF,t_70)
```python
conda install cloudpickle
```
最后，spyder4.0如果启动出现“crashed during last session”，可能是kite的问题，卸载kite可以解决。

```python
#设置关键文件的路径
path_template =r'./templates'
path_xlsx =r'./templates/template_x.xlsx'
path_docx1 =r'./templates/template_d_entrust.docx'
path_docx2 =r'./templates/template_d_communication.docx'

#因为存在nan值，默认会转换成float型，手机号输出会带小数点，这里指定'telephone'列为Int64型,pd在0.24以后版本已经可以将含有nan值的数组保存为整型。
dtype_dic= {'telephone': 'Int64','instrument_numbers': 'Int64' }
#读取excel中的委托信息
df = pd.read_excel (path_xlsx, sheet_name=1, header=0,index_col=None, na_values = [ 'NA' ], dtype = dtype_dic)

```
# 委托信息整理

```python
#获取最后一行信息的索引
lastrow=df.index[-1]

#对'instrument_numbers'列因合并单元格的产生的nan值进行填充
df['instrument_numbers'].fillna(method='pad',inplace=True)

#对接受日期的年月日信息进行分列，方便后续填入到文档中
df['YY'], df['MM'] , df['DD']= df['接收日期'].str.split('.', 2).str
```

# 委托信息替换

```python
#读取模板文档
tpl_1 = DocxTemplate(path_docx1)
tpl_2 = DocxTemplate(path_docx2)

#替换word中的变量
#字典中的key为变量名，value为要替换的值
context = { 
    'Delegate_numbers':df.Delegate_numbers[lastrow],
    'client':df.client[lastrow],
    'addr':df.addr[lastrow],
    'name':df.name[lastrow],
    'telephone':df.telephone[lastrow],
    'instrument_name':df.instrument_name[lastrow],
    'instrument_produce':df.instrument_produce[lastrow],
    'instrument_numbers':df.instrument_numbers[lastrow],
    'instrument_model':df.instrument_model[lastrow],
    'instrument_sn':df.instrument_sn[lastrow],
    'YY':df.YY[lastrow],
    'MM':df.MM[lastrow],
    'DD':df.DD[lastrow],
    'standard_multi':R(std),#template提供了5种方式对字符进行转义,见https://docxtpl.readthedocs.io/en/latest/index.html
    'standard_solo':std1+','+std2
    
}

tpl_1.render(context, autoescape=True)
path_save1='委托单_{}_{}.docx'.format(df.loc[lastrow,'Delegate_numbers'],df.loc[lastrow,'client'])
tpl_1.save(path_save1)

tpl_2.render(context, autoescape=True)
path_save2='与客户沟通的记录及评审表_{}_{}.docx'.format(df.loc[lastrow,'Delegate_numbers'],df.loc[lastrow,'client'])
tpl_2.save(path_save2)
```


# 完整代码

```python
在这里插入代码片
```

# 参考文献

[[1]:[Python模板引擎——jinja2的基本用法集锦]](https://www.jianshu.com/p/3bd05fc58776)  
[2]:[利用python批量出报告](https://www.capallen.top/2019/%E5%88%A9%E7%94%A8python%E6%89%B9%E9%87%8F%E5%86%99%E6%8A%A5%E5%91%8A)  
[3]:[Python办公自动化|excel读取和写入](https://www.capallen.top/2019/%E5%88%A9%E7%94%A8python%E6%89%B9%E9%87%8F%E5%86%99%E6%8A%A5%E5%91%8A)  
[4]:[Python办公自动化|批量word报告生成工具](https://www.capallen.top/2019/%E5%88%A9%E7%94%A8python%E6%89%B9%E9%87%8F%E5%86%99%E6%8A%A5%E5%91%8A)  
[5]:[python-docx|template 操作word文档](https://www.capallen.top/2019/%E5%88%A9%E7%94%A8python%E6%89%B9%E9%87%8F%E5%86%99%E6%8A%A5%E5%91%8A)  
[6]:[超简单Python将Excel的指定数据插入到docx模板并生成](https://blog.csdn.net/weixin_41133061/article/details/88543432)  
[7]:[Excel信息批量替换Word模板生成新文件](https://blog.csdn.net/chen8782186/article/details/98784005?depth_1-utm_source=distribute.pc_relevant.none-task&utm_source=distribute.pc_relevant.none-task)  
[8]:[Pandas查找缺失值的位置，并返回缺失值行号以及列号](https://blog.csdn.net/u010924297/article/details/80060229)  
[9]:[基于docxtpl的自动化报告生成(基于word模板)](https://blog.csdn.net/yycoolsam/article/details/103255271?ops_request_misc=%257B%2522request%255Fid%2522%253A%2522158488225319724811860476%2522%252C%2522scm%2522%253A%252220140713.130056874..%2522%257D&request_id=158488225319724811860476&biz_id=0&utm_source=distribute.pc_search_result.none-task)  




[另外这一篇作者只使用了python-docx库也达到了批量出具报告的效果](https://cloud.tencent.com/developer/article/1573184)
