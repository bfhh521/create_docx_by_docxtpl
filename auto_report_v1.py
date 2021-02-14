import pandas as pd 
from pathlib import Path
from jinja2 import Environment,FileSystemLoader
from docxtpl import *
import os,sys

def mkdir(path):
 
	folder = os.path.exists(path)
 
	if not folder:                   #判断是否存在文件夹如果不存在则创建为文件夹
		os.makedirs(path)            #makedirs 创建文件时如果路径不存在会创建这个路径
		print ('生成当前报告文件夹,文件夹为\'./{}\''.format(df_ms.loc[lastrow,'Delegate_numbers']))
	else:
		print ('当前报告文件夹已存在,不再新建,文件夹为\'./{}\''.format(df_ms.loc[lastrow,'Delegate_numbers']))
        
def entrust_data_output():        
    #----------------------------------------数据输出------------------------------------------
    #读取模板文档
    tpl_1 = DocxTemplate(path_entrust)
    tpl_2 = DocxTemplate(path_communication)
    
    #替换word中的变量
    #字典中的key为变量名，value为要替换的值
    context = { 
        'Delegate_numbers':df_ms.Delegate_numbers[lastrow],
        'client':df_ms.client[lastrow],
        'addr':df_ms.addr[lastrow],
        'name':df_ms.name[lastrow],
        'telephone':df_ms.telephone[lastrow],
        'instrument_name':df_ms.instrument_name[lastrow],
        'instrument_produce':df_ms.instrument_produce[lastrow],
        'instrument_numbers':df_ms.instrument_numbers[lastrow],
        'instrument_model':df_ms.instrument_model[lastrow],
        'instrument_sn':df_ms.instrument_sn[lastrow],
        'YY':df_ms.YY_re[lastrow],
        'MM':df_ms.MM_re[lastrow],
        'DD':df_ms.DD_re[lastrow],
        'standard_multi':R(std),#template提供了5种方式对字符进行转义,见https://docxtpl.readthedocs.io/en/latest/index.html
        'standard_solo':std1+','+std2
        
    }
    
    #生成委托单并保存
    tpl_1.render(context, autoescape=True)
    path_save1='委托单_{}_{}.docx'.format(df_ms.loc[lastrow,'Delegate_numbers'],df_ms.loc[lastrow,'client'])
    path_save1=os.path.join(new_folder, path_save1)
    tpl_1.save(path_save1)
    
    #生成沟通记录并保存
    tpl_2.render(context, autoescape=True)
    path_save2='与客户沟通的记录及评审表_{}_{}.docx'.format(df_ms.loc[lastrow,'Delegate_numbers'],df_ms.loc[lastrow,'client'])
    path_save2=os.path.join(new_folder, path_save2)
    tpl_2.save(path_save2)
    
    #对文件是否成功生成进行检查
    print('==============文件生成检查==============\n')
    if Path(path_save1).is_file() != True:
        print('注意，委托单文件未成功生成')
    elif Path(path_save2).is_file() != True:
        print('注意，沟通记录文件未成功生成')           
    else:
        print('委托单、沟通记录文件均已生成,文件名为\'委托单/与客户沟通的记录及评审表_{}_{}\'。'.format(df_ms.loc[lastrow,'Delegate_numbers'],df_ms.loc[lastrow,'client']))

def report_data_output(): 
    
    #读取data数据
    df_data = pd.read_excel (path_data, sheet_name=0, header=0,index_col=None, na_values = [ 'NA' ])

    tpl=DocxTemplate(path_report)#读取模板
    
    Ue=5.1#设定电场不确定度下限
    Um=3.1#设定磁场不确定度下限
        
    '''生成的跟table一样的列表，但不知道输入context后，render出来只有一行空白数据
    k=[]
    for i in range(1,10):
        f=[''.join ('data%d' % (x)) for x in range(1,6)]#批量生成变量名的列表
        g=[i,df_data.标准值[i-1],df_data.指示值[i-1],df_data.修正值[i-1],5.1 if df_data.Urel[i-1] <= 5.1 else df_data.Urel[i-1] ]
        h=[dict(zip(f,g))]#批量生成变量名的字典并转换成列表
        k.append(h)
    '''   
    
    table_1 = \
             [{'data1': i,
               'data2': '-' if format('%.2f' %df_data.标准值[i-1]) == 'nan' else format('%.2f' %df_data.标准值[i-1]),
               'data3': '-' if format('%.2f' %df_data.指示值[i-1]) == 'nan' else format('%.2f' %df_data.指示值[i-1]),
               'data4': '-' if format('%.2f' %df_data.修正值[i-1]) == 'nan' else '0' if df_data.修正值[i-1] == 0 else format('%.2f' %df_data.修正值[i-1]),
               'data5': format('%.1f' %Ue) if df_data.Urel[i-1] <= Ue else format('%.1f' %df_data.Urel[i-1]), }
              for i in range(1,13)]
             
    table_2 = \
             [{'data1': i,
               'data2': '-' if format('%.2f' %df_data.标准值[i+13]) == 'nan' else format('%.2f' %df_data.标准值[i+13]),
               'data3': '-' if format('%.2f' %df_data.指示值[i+13]) == 'nan' else format('%.2f' %df_data.指示值[i+13]),
               'data4': '-' if format('%.2f' %df_data.修正值[i+13]) == 'nan' else '0' if df_data.修正值[i+13] == 0 else format('%.2f' %df_data.修正值[i+13]),
               'data5': format('%.1f' %Ue) if df_data.Urel[i+13] <= Ue else format('%.1f' %df_data.Urel[i+13]), }
              for i in range(1,13)]
             
    table_3 = \
             [{'data1': i,
               'data2': '-' if format('%.2f' %df_data.标准值[i+27]) == 'nan' else format('%.2f' %df_data.标准值[i+27]),
               'data3': '-' if format('%.2f' %df_data.指示值[i+27]) == 'nan' else format('%.2f' %df_data.指示值[i+27]),
               'data4': '-' if format('%.2f' %df_data.修正值[i+27]) == 'nan' else '0' if df_data.修正值[i+27] == 0 else format('%.2f' %df_data.修正值[i+27]),
               'data5': format('%.1f' %Ue) if df_data.Urel[i+27] <= Ue else format('%.1f' %df_data.Urel[i+27]), }
              for i in range(1,13)]
             
    table_4 = \
             [{'data1': i,
               'data2': '-' if format('%.2f' %df_data.标准值[i+41]) == 'nan' else format('%.2f' %df_data.标准值[i+41]),
               'data3': '-' if format('%.2f' %df_data.指示值[i+41]) == 'nan' else format('%.2f' %df_data.指示值[i+41]),
               'data4': '-' if format('%.2f' %df_data.修正值[i+41]) == 'nan' else '0' if df_data.修正值[i+41] == 0 else format('%.2f' %df_data.修正值[i+41]),
               'data5': format('%.1f' %Um) if df_data.Urel[i+41] <= Um else format('%.1f' %df_data.Urel[i+41]), }
              for i in range(1,13)]
              
    table_5 = \
             [{'data1': i,
               'data2': '-' if format('%.2f' %df_data.标准值[i+55]) == 'nan' else format('%.2f' %df_data.标准值[i+55]),
               'data3': '-' if format('%.2f' %df_data.指示值[i+55]) == 'nan' else format('%.2f' %df_data.指示值[i+55]),
               'data4': '-' if format('%.2f' %df_data.修正值[i+55]) == 'nan' else '0' if df_data.修正值[i+55] == 0 else format('%.2f' %df_data.修正值[i+55]),
               'data5': format('%.1f' %Um) if df_data.Urel[i+55] <= Um else format('%.1f' %df_data.Urel[i+55]), }
              for i in range(1,13)]
             
    table_6 = \
             [{'data1': i,
               'data2': '-' if format('%.2f' %df_data.标准值[i+69]) == 'nan' else format('%.2f' %df_data.标准值[i+69]),
               'data3': '-' if format('%.2f' %df_data.指示值[i+69]) == 'nan' else format('%.2f' %df_data.指示值[i+69]),
               'data4': '-' if format('%.2f' %df_data.修正值[i+69]) == 'nan' else '0' if df_data.修正值[i+69] == 0 else format('%.2f' %df_data.修正值[i+69]),
               'data5': format('%.1f' %Um) if df_data.Urel[i+69] <= Um else format('%.1f' %df_data.Urel[i+69]), }
              for i in range(1,13)]
              
    context = {
        'Delegate_numbers':df_ms.Delegate_numbers[lastrow],
        'client':df_ms.client[lastrow],
        'instrument_name':df_ms.instrument_name[lastrow],
        'instrument_produce':df_ms.instrument_produce[lastrow],
        'instrument_numbers':df_ms.instrument_numbers[lastrow],
        'instrument_model':df_ms.instrument_model[lastrow],
        'instrument_sn':df_ms.instrument_sn[lastrow],
        'YY':df_ms.YY_ca[lastrow],
        'MM':df_ms.MM_ca[lastrow],
        'DD':df_ms.DD_ca[lastrow],
        'Temp':df_ms.Temp[lastrow],
        'humidity':df_ms.humidity[lastrow],
        'Self_E':df_ms.Self_E[lastrow],
        'Self_H':df_ms.Self_H[lastrow],
        'table_1': table_1,
        'table_2': table_2,
        'table_3': table_3,
        'table_4': table_4,
        'table_5': table_5,
        'table_6': table_6
        }
    
    
    #生成报告并保存
    tpl.render(context) 
    path_save='report_{}.docx'.format(df_ms.loc[lastrow,'Delegate_numbers'])
    path_save=os.path.join(new_folder, path_save)
    tpl.save(path_save) 

    #对文件是否成功生成进行检查
    print('==============报告生成检查==============\n')
    if Path(path_save).is_file() != True:
        print('注意，报告未成功生成')         
    else:
        print('报告已生成，文件名为\'report_{}\'。'.format(df_ms.loc[lastrow,'Delegate_numbers']))     

def data_cleansing(): 
    global df_ms,std,std1,std2,new_folder
    #----------------------------------------整理信息------------------------------------------
        
    #对获取的最后一行顾客信息是否缺失进行检查
    print('==============顾客信息检查==============\n')
    if df_ms.loc[lastrow].count() != df_ms.shape[1]:
        loc = df_ms.loc[lastrow][df_ms.loc[lastrow].isnull().values==True].index.tolist()
        print('注意，报告{}在{}信息有缺失值'.format(df_ms.loc[lastrow,'Delegate_numbers'],loc))
    else:
        new_folder = './{}'.format(df_ms.loc[lastrow,'Delegate_numbers'])
        mkdir(new_folder)             #调用函数新建当前报告文件夹
        print('顾客信息检查通过。')
    
    #对接收日期,校准日期的年月日信息进行分列，方便后续填入到文档中
    df_ms=df_ms.join(df_ms['receipt date'].str.split('.', 2, expand=True).rename(columns={0:'YY_re', 1:'MM_re', 2:'DD_re'}))
    df_ms=df_ms.join(df_ms['calibration date'].str.split('.', 2, expand=True).rename(columns={0:'YY_ca', 1:'MM_ca', 2:'DD_ca'}))
    
    #针对工频和直流项目所使用的不同标准进行写入
    if df_ms.instrument_name[lastrow]=='工频场强计':
        std1='DL/T 988—2005《高压送电线路、变电站工频电磁场测量方法》附录A、附录B'
        std2='JJG 1049-2009《弱磁场交变磁强计检定规程》'
        std='{}\n{}'.format(std1,std2)    
    elif df_ms.instrument_name[lastrow]=='直流合成场强计':
        std1='DL/T 1089-2008 《直流换流站与线路合成场强、离子流密度测量方法》附录A'
        std2='直流合成场强测量仪（场磨）校准方法'
        std='{}\n{}'.format(std1,std2)
    else:
        std=''        

#----------------------------------------读取文档------------------------------------------

#设置关键文件的路径
path_template =r'./templates'
path_message =r'./校准流程2020.xlsx'
path_entrust =r'./templates/entrust_tpl.docx'
path_communication =r'./templates/communication_tpl.docx'
path_report =r'./templates/report_tpl.docx'
path_data =r'./data.xlsx'


#对默认路径的模板文件是否缺失进行检查
print('==============模板文件检查==============\n')
if Path(path_template).is_dir() != True:
    print('注意，模板文件夹缺失，程序已终止')
    sys.exit(0)
elif Path(path_message).is_file() != True:
    print('注意，校准流程的xlsx文件缺失，程序已终止')
    sys.exit(0)
elif Path(path_entrust).is_file() != True:
    print('注意，委托单的模板文件缺失，程序已终止')
    sys.exit(0)
elif Path(path_communication).is_file() != True:
    print('注意，沟通记录的模板文件缺失，程序已终止')  
    sys.exit(0)
elif Path(path_report).is_file() != True:
    print('注意，报告的模板文件缺失，程序已终止') 
    sys.exit(0)              
else:
    print('模板文件检查通过。')

#因为存在nan值，默认会转换成float型，手机号输出会带小数点，这里指定'telephone'列为Int64型,pd在0.24以后版本已经可以将含有nan值的数组保存为整型。
dtype_dic= {'telephone': 'Int64','instrument_numbers': 'Int64' }
#读取excel中的委托信息
df_ms = pd.read_excel (path_message, sheet_name=1, header=0,index_col=None, na_values = [ 'NA' ], dtype = dtype_dic)
#获取最后一行信息的索引
lastrow=df_ms.index[1]
#lastrow=df_ms.index[-1]
#对'instrument_numbers'列因合并单元格的产生的nan值进行填充
df_ms['instrument_numbers'].fillna(method='pad',inplace=True)

    

if Path(path_data).is_file() != True:
    print('注意，报告的数据文件缺失，本次将不会生成报告') 
    data_cleansing()    
    entrust_data_output()
else:
    print('报告的数据已导入，即将生成报告') 
    data_cleansing()    
    report_data_output()





    


