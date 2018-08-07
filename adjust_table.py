import  sys
import xlwt
from functools import reduce
import re
import json
import subprocess,os
import shutil
import hashlib
from  hashlib import md5
import PyPDF2

table_header=[['名称、地名代码','水平范围','备注'],
                            ['名称、水平范围、垂直范围', '提供服务的单位', '呼号', '工作频率(*表示备用)', '服务时间', '备注'],
                            ['名称、范围、高度','提供服务的单位','话呼','工作频率、时间','备注'],
                            ['航路、航线代号、导航点名称、坐标','磁航迹距离(千米/海里)','最低飞行高度(米)','宽度(千米)','巡航高度层方向','管制单位'],
                            ['名称代码','坐标','航路'],
                          ['未出席董事职务','未出席董事姓名', '未出席董事的原因说明','被委托人姓名'],
                          ['董事会秘书 ', '证券事务代表'],
                          ['释义项','指','释义内容'],
                          ['主要资产','重大变化说明'],
                          [' ','2017','2016','2015'],
                          ['2017','2016','2015'],
                          ['项目 ','房屋及建筑物',' 专用设备 ','运输工具 ','电子及其他设备 ','合计'],
                          ['类别 ','折旧方法 ','折旧年限(年) ','残值率(%) ','年折旧率(%) '],
                          ['项目',' 土地使用权 ','专利权及非专利技术',' 管理软件 ','商标',' 特许经营权 ','合计 '],
                          ['项目',' 摊销年限(年) '],
                          ['项目 ','本期数 ','上年同期数',' 计入本期非经常性损益的金额'],
                          ['项目',' 房屋及建筑物',' 机器设备 ','运输工具',' 电子设备、器具及家具',' 合计 '],
                          ['类别',' 折旧方法',' 折旧年限（年）',' 残值率 ','年折旧率 '],
                          ['项目 土地使用权',' 非专利技术 ','软件 ','其他',' 合计 '],
                          ['类别','摊销年限(年)',' 年摊销率(%) '],
                          ['项目',' 本期发生额',' 上期发生额',' 计入当期非经常性损益的金额 '],
                          ['联系人和联系方式',' 董事会秘书',' 证券事务代表 '],
                          [' ','2017年 ','2016年',' 本年比上年增减 ','2015年 '],
                          ['第一季度',' 第二季度 ','第三季度 ','第四季度 '],
                          ['产品名称','营业收入','营业利润','毛利率','营业收入比上年同期增减','营业利润比上年同期增减','毛利率比上年同期增减'],
                          ['资产的具体内容','形成原因','资产规模','所在地','运营模式','保障资产安全性的控制措施','收益状况','境外资产占公司净资产的比重','是否存在重大减值风险'],
                            ]
"""
tabula两种输出模式：lattice和stream。
目前代码是根据tabula的lattice模式输出做调整。输出是json文件，文件保留了每个单元格的一些位置信息。
从结果输出统计来看，如果pdf中有黑色背景的表格，stream输出结果的完整性更好，但是缺少结构信息，没办法判断表格的边界在哪里。
所以综合来看，应该用一种新的模式结合现有的两种模式的优缺点。
关于边界的概念，有个功能需要完善的地方：
当某个表格内容跨多页时，怎么在输出时合并成一个表格？
----
还有，现有的调整基础可控制性太弱了，最有效的方式应该在tabula原项目上查找原因。
"""
tabula_path =  'tabula/tabula-1.0.1-jar-with-dependencies.jar'
#一页一个sheet
def write_page_to_excel_by_sheet(filename,tables):

    pages=len(tables)
    book=xlwt.Workbook()
    text_tables = output_data(tables)
    # print(len(text_tables))
    for i in range(0,len(text_tables)):
        
        table=text_tables[i]
        data=[]
        if has_data(table):
            data=table['data']
        else:
            data=table
            
        sheetN = book.add_sheet('第%d个表格' % (i+1) ,cell_overwrite_ok=True) #创建sheet
        cell_style = xlwt.easyxf('pattern: pattern solid,fore_colour white; alignment: wrap 1;') 
        for j in range(0,len(data)):
            
            row=data[j]   
            
            for k in range(0,len(row)):
            
                cell=row[k]
                text=''
                if type(cell) == type(''):
                    text=cell
                    text=text.replace('\r','\\r\r')
                    text=text.replace('\n','\\n\n')
                else:
                    text=str(cell)
                sheetN.write(j,k,text,cell_style)
                sheetN.col(k).width=5000
           
    book.save(filename) 
    
def write_tables_to_excel(filename,tables):

    book=xlwt.Workbook()
    sheet1 = book.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet
    style = xlwt.XFStyle() # 初始化样式
    aligm=xlwt.Alignment()
    aligm.wrap=True
    style.alignment=aligm
    
    xlwt.add_palette_colour("custom_colour", 0x21)
    book.set_colour_RGB(0x21, 220, 215, 223)
    
    total_rows=0
    for data in tables:
        for j in range(0,len(data)):
            row=data[j]   
            if j%2==0:
                cell_style = xlwt.easyxf('pattern: pattern solid,fore_colour custom_colour; alignment: wrap 1;')  
            else:
                cell_style = xlwt.easyxf('pattern: pattern solid,fore_colour white; alignment: wrap 1;')             
            # style.pattern = pattern
            
            for k in range(0,len(row)):
            
                text=row[k]
                
                sheet1.write(j+total_rows,k,text,cell_style)
                sheet1.col(k).width=5000
                # sheet1.col(k).width=500
                # Tools.output_file('adjust_table.txt',o_text,new_line=(k==len(row)-1))
                # print('%s' % str(cell_width),end=' '*4)
        
        total_rows+=len(data)
           
    book.save(filename) #保存文件   

    
def write_excel(filename,data):
    print('正在将数据写入到%s文件...'%filename)
    f = xlwt.Workbook() #创建工作簿
    '''
    创建第一个sheet:
    sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet
    style = xlwt.XFStyle() # 初始化样式
    aligm=xlwt.Alignment()
    aligm.wrap=True
    style.alignment=aligm
    # Tools.output_file('temp_of_excel_object.txt',dir(sheet1))

    # print(sheet1.col_width(0))
    #生成第一行
    # for i in range(0,len(row0)):
        # sheet1.write(0,i,row0[i],set_style('Times New Roman',220,True))
    
    # table_height=table['height']
    # table_width=table['width']
    # table_top=table['top']
    # table_left=table['left']
    # data=table['data']
    row_len=0
    # print('cell width:')
    for j in range(0,len(data)):
        row=data[j]
        # if row_len!=0 and len(row)!=row_len:
            # print('两行列数不同')
        row_len=len(row)       
        # sheet1.row(j).height=200
        # print(row)
        for k in range(0,len(row)):
            cell=row[k]
            cell_height=cell['height']
            cell_width=cell['width']
            cell_left=cell['left']
            cell_top=cell['top']
            text=cell['text']
            # if len(text)>5:
                # simple_text=text[0:5]+'...'
            # else:
                # simple_text=text
            xy_text=''
            if 'x' in cell.keys() and 'y' in cell.keys():
                xy_text='x:%s,y:%s' %(cell['x'],cell['y'])
            o_text='top:%s\r\nleft:%s\r\nwidth:%s\r\nheight:%s\r\n%s\n%s'%(str(cell_top),str(cell_left),str(cell_width),str(cell_height),xy_text,text)
            sheet1.write(j,k,o_text,style)
            sheet1.col(k).width=7000
            # sheet1.col(k).width=500
            # Tools.output_file('adjust_table.txt',o_text,new_line=(k==len(row)-1))
            # print('%s' % str(cell_width),end=' '*4)

           
    f.save(filename) #保存文件
    print('%s写入成功'%filename)

# write_excel('adjust_table_example_data.xls',test_table)

def compare(one,two):
    wucha=0.5
    if one == two or abs(one-two)<=wucha:
        #0.5是误差值
        #记录下比较误差值 判断误差值准确范围
        # Tools.output_file('adjust_table_TEST_wuchazhi.txt',abs(one-two),clear=False)
        return 0
    if one-two>wucha:
        return 1
    return -1


    
def adjust(extract_table):

        
    table=extract_table.copy()
    header_index=find_header(table)
    if header_index<0:
        # print('检测到无标题栏的表格数据.')
        return None
    data=table['data']
    header=data[header_index]  
    
    table_height=table['height']
    table_width=table['width']
    table_top=table['top']
    table_left=table['left']
    row_len=0
    
    mark_data=data[:]
    row_obj={'height':0.0,'top':0.0}
    row_info=[]
    
    #1.查找行边界
    for r in range(0,len(mark_data)):
        
        row=mark_data[r]
        row_height=max([obj['height'] for obj in row])
        row_top=max([obj['top'] for obj in row])
        
        if compare(row_height+row_top,row_obj['height']+row_obj['top'])>0:
            row_obj['height']=row_height
            row_obj['top']=row_top
            row_info.append(row_obj.copy())
            
    # print(row_info)
    #2.列边界从标题栏中截取
    #3.标记
    for r in range(0,len(mark_data)):
        
        row=mark_data[r]
        for c in range(0,len(row)):
            cell=row[c]
            if is_empty_cell(cell):
                continue
            else:
                y=-1
                x=-1
              
                # print('mark cell text:%s  left:%s  width:%s' % (cell['text'],str(cell['left']),str(cell['width'])))
                #3.1查找row坐标
                for i in range(0,len(row_info)):
                    info=row_info[i]
                    if compare(cell['top']+cell['height'],info['top']+info['height'])<=0  and compare(cell['top']+cell['height'],info['top'])>0:
                        y=i
                        break
                
                #3.2查找col坐标
                for j in range(0,len(header)):
                    if compare(cell['left']+cell['width'],header[j]['left']+header[j]['width'])<=0:
                        if compare(cell['left']+cell['width'],header[j]['left'])>0:
                            x=j
                            # print('1-x:%d/%d' % (x,j))
                            #跨列标记
                            if compare(cell['left'],header[j]['left'])<0:
                                start=0
                                # end=0
                                for k in range(0,len(header)):
                                    if compare(cell['left'],header[k]['left'])==0:
                                        start=k
                                    # if compare(cell['left']+cell['width'],header[k]['left']+header[k]['width'])<=0 and compare(cell['left']+cell['width'],header[k]['left'])>0:
                                        # end=k
                                x=start
                            break 
                            
                # print('2-x:%d' % x)
            mark(cell,x,y)
            # print(cell)
    
    #4.移动
    new_table=[]
    
    for r in range(0,len(row_info)):
        new_row=[]
        for c in range(0,len(header)):
            
            #查找相同标记的空白格
            one_piece=[]
            for i in range(0,len(data)):
                row=data[i]
                for j in range(0,len(header)):
                    cell=row[j]
                    # print('find xy cell:' % cell)
                    if is_empty_cell(cell):
                        continue
                    else:
                        if cell['x']==c and cell['y'] ==r:
                            find_result=-1
                            for ocell in one_piece:
                                find_result=ocell['text'].find(cell['text'])
                                if  find_result>=0:
                                    # Tools.output_file('debug_the_same_cell.txt','ocell text: %s' % ocell['text']+'\ncell text:%s' % cell['text']+'\n')
                                    break

                            if find_result<0:
                                one_piece.append(cell)
            
            #拼接
            if len(one_piece)>0:
                # if len(one_piece)
                one_piece=merge_cells(one_piece)
                new_row.append(one_piece[0])
        new_table.append(new_row)
        table['data']=new_table
    return table

def mark(cell,x,y):
    cell['x']=x
    cell['y']=y
  
def merge_cells(cells):
    #去重複元素的高深莫测写法~~
    func=lambda x,y:x if y in x else x+[y]
    left_list=reduce(func,[[],]+[obj['left'] for obj in cells])
    # print(left_list)

    temp_cells=sorted(cells,key=lambda obj:obj['top'])
    
    col_merge_symbol='\r'
    for left in left_list:
        col_min_obj=None
        i=0
        col_merge_text=''
        # col_merge_height=0
        while i<len(temp_cells):
            obj=temp_cells[i]
            step=1
            if obj['left'] == left:
                if col_min_obj is None:
                    col_min_obj=obj
                    col_merge_text=obj['text']
                    # col_merge_height=obj['height']
                else:
                    if compare(obj['top'],col_min_obj['top'])<0:
                        print('列顺序不对')
                    else:    
                        if col_merge_text =='':
                            col_merge_text = obj['text']
                        elif obj['text'] !='':
                            col_merge_text=col_merge_text+col_merge_symbol+obj['text']
                        step=0
                        col_min_obj['height']+=obj['height']
                        temp_cells.pop(i)
                        
            i+=step
        if col_min_obj:
            col_min_obj['text']=col_merge_text
            
    row_min_obj=[obj for obj in  temp_cells if obj['top'] == min([obj['top'] for obj in temp_cells])][0]
    # print(row_min_obj)
    row_merge_text=row_min_obj['text']
    row_merge_width=row_min_obj['width']
    row_min_left=row_min_obj['left']
    j=0
    while j<len(temp_cells):
        step=1
    # for obj in temp_cells:
        obj=temp_cells[j]
        # if compare(row_min_obj['top'],obj['top'])!=0:  
            # print('行top值不同,row_min_obj[top]:%s/obj[top]:%s' % (str(row_min_obj['top']),str(obj['top'])))
        
        if obj != row_min_obj:
            # print(obj)
            row_merge_text+=obj['text']
            row_merge_width+=obj['width']
            if compare(row_min_obj['left'],obj['left'])>0:
                row_min_left=obj['left']
            temp_cells.pop(j)
            step=0
        j+=step
        
    row_min_obj['width']=row_merge_width
    row_min_obj['left']=row_min_left
    row_min_obj['text']=row_merge_text
    return temp_cells
    
def is_empty_row(row):
    for cell in row:
        if not is_empty_cell(cell):
            return False
    return True 
    
def is_empty_cell(cell):
    cell_height=cell['height']
    cell_width=cell['width']
    cell_left=cell['left']
    cell_top=cell['top']
    cell_text=cell['text']
    if cell_top !=0 or cell_left !=0 or cell_width!=0 or cell_height!=0 or cell_text!='':
        return False
    return True
     
def has_data(table):
    if type(table) == type({}):
        if 'data' in table.keys():
            return True
    return False   
 
def adjust_tables(tables):
    adjust_data=[]
    for table_data in tables:
        result=adjust(table_data)
        if result:
            adjust_data.append(result)
    return adjust_data
    
def output_data(tables):
    text_data=[]
    for table in tables:
        text_data.append(output_table_data(table))
    return text_data

def output_table_data(table):
    new_table=[]
    data=[]
    # print(has_data(table))
    if has_data(table):
        data=table['data']
    else:
        data=table
    for row in data:
        new_row=[]
        for cell in row:
            # print(cell)
            new_row.append(cell['text'])
        new_table.append(new_row)
    return new_table
    
def get_pages(file_name):    
    pdf_file = PyPDF2.PdfFileReader(file_name)
    pages = pdf_file.getNumPages()
    return pages
    
def is_header(row):
    row_text=get_row_text(row)
    for header in table_header:
        # print('key:%s,value:%s' % (key,header))
        header_text=''.join(header)
        if row_text == header_text:
            return True
    return False     

def find_header(table):
    data=table['data']
    for i in range(0,len(data)):
            row=data[i]
            if is_header(row):    
                return i
    return -1
    
def has_header(table):
        
        if type(table) == type({}):
            if 'data' in table.keys():
                data=table['data']
            else:
                data=table
        else:
            data=table
        for i in range(0,len(data)):
            row=data[i]
            if is_header(row):  
                return True
        return False
        
def get_row_text(row):
    return ''.join([obj['text'] for obj in row]).replace('\r','').strip()

def read_file(filename,op='r',json_data=True):
    data=[]
    with open(filename,op,encoding='gb18030') as f:
        if json_data:
            data=json.load(f)
        else:
            data=f.read()
    return data
    
def generation_page_file(target_file_name,extract_mode,output_file_type,page):

    # temp_file_name=get_page_file_name(target_file_name,extract_mode,page)
    new_file_name = os.path.splitext(os.path.split(target_file_name)[1])[0] + '-page-{}.json'.format(page)
    page_file_name = os.path.join('output/', new_file_name)
    # print('第%d页：%s' % (page,page_file_name))
    java_cm='java -jar %s %s -%s -f %s -p %d -o %s' %(tabula_path,target_file_name,extract_mode,output_file_type,page, page_file_name)
    # print('java cmd:%s' % java_cm)
    exec_status,exec_output=subprocess.getstatusoutput(java_cm)
    table=[]
    if exec_status==0:
        if os.path.exists(page_file_name):
            table=read_file(page_file_name)
            # print('generation_page_file/read_file:%s' % table)
            os.remove(page_file_name)
    else:
        if exec_output.find('Page number does not exist'):
            exec_status=-100
    return table,exec_status
    
def correction_data(lattice_data,stream_data):
    
    lattice_table_str=json.dumps(lattice_data)
    stream_table_str=json.dumps(stream_data)
    
    corr_re=re.compile(r'N[0-9]{6}E[0-9]{6}(?=[^0-9])')
    bad_data=list(set(corr_re.findall(lattice_table_str)))
    # print('-'*i+repr(i))
    # print('bad_data length:%d' % len(bad_data))
    if len(bad_data)>0:
        for k in bad_data:
            good_data=re.findall(k+'[0-9]',stream_table_str)
            if good_data:
                if len(set(good_data))!=1:
                    print('修复数据出现多个选择！%s' % str(good_data))
                print('%s将替换为：%s' %(k,list(set(good_data))[0]))
                lattice_table_str=lattice_table_str.replace(k+'(?=[^0-9])',list(set(good_data))[0])
                
        # corr_data=corr_re.findall(lattice_table_str)
        # print('corr_data length:%d' % len(corr_data))
        # if len(corr_data)==0:
            # print('(%d)页修复完成.'% i )
            
    return json.loads(lattice_table_str)
    
def find_param(param_name, default=None):

    param_prefix =param_name + '='
    for a in range(2, len(sys.argv)):
        arg = sys.argv[a]
        if arg.find(param_prefix) >=0:
            index = arg.find(param_prefix)
            value = arg[index+len(param_prefix):]
            print('param: %s = %s' % (param_name, value))
            return value
    return default        
    
def repair():

    if len(sys.argv)>1:
        target=sys.argv[1]
        # print(target)
        
        if os.path.exists(target):
        
            pdf_files=[]
            # tabula_path='tabula/tabula-1.0.1-jar-with-dependencies.jar'
            file_suffix = '.pdf'
            if os.path.isfile(target):
                # print('a file.')
                if os.path.splitext(target)[1] == file_suffix:
                    pdf_files.append(target)
                    
            elif os.path.isdir(target):
                # print('is dir.')

                for file in os.listdir(target):
                    file_path=os.path.join(target,file)
                    # print(file)
                    # print(os.path.splitext(file_path)[1] == 'pdf')
                    # print('读取到的pdf文件有：')
                    if os.path.isfile(file_path) and os.path.splitext(file_path)[1] == file_suffix:
                        pdf_files.append(file_path)
                        
            #分页检查开关
            truish_words = ['yes', 'true', '1', 'y', 'ok']
            need_check = find_param('check', 'True').lower() in truish_words
            print('need check') if need_check else print('only output all')
            
            lattice_mode = ['lattice', 'l', 'r', 'spreadsheet']
            stream_mode =['stream', 's', 'n']
            extract_mode = find_param('mode', 'r')
            if not extract_mode or extract_mode not in lattice_mode+stream_mode:
                extract_mode = 'r'
                
            for file in pdf_files:
                 
                #分两种形式输出调整后的表格：
                #1.单页分开，一页一个文档。目的是为了检验和对照调整的正确性。
                #   是否启用单页输出由命令行第二个参数值小写是否在数组['yes', 'true', '1', 'y', 'ok']中决定，默认是True
                #2.整个文件输出到一个文档
                #   始终启用
                
                file_name = os.path.splitext(os.path.split(file)[1])[0]
                print('original file name :%s' % file_name )
                file_name_hash=hashlib.md5(file_name.encode('utf-8')).hexdigest()
                temp_dir='temp/'
                if not os.path.exists(temp_dir):
                    os.mkdir(temp_dir)
                
                adjusting_file_name=os.path.join(temp_dir, file_name_hash+file_suffix)
                # print('file_name_hash:%s' % file_name)
                # if os.path.exists(file_name):
                    # os.remove(file_name)
                shutil.copy2(file,adjusting_file_name)
                
                pages = get_pages(adjusting_file_name)
                print('check pages:%d' % pages)
                
                # if os.path.exists(file_name):
                    # print('%s复制完成。' % file_name)
                    
                file_type='JSON'
                temp_file_suffix=''
                if file_type == 'JSON':
                    temp_file_suffix='.json'
                elif file_type == 'CSV':
                    temp_file_suffix='.csv'
                 
                 
                output_dir='output/'
                if not os.path.exists(output_dir):
                    os.mkdir(output_dir)
                    
                # print('file :%s' % file)
                adjusting_all_file_name = os.path.join(temp_dir, file_name_hash+'-all'+temp_file_suffix)
                java_cm=u'java -jar %s %s -%s -f %s -p %s -o %s' %(tabula_path,adjusting_file_name,extract_mode,file_type,'all',adjusting_all_file_name)
                print('all java cmd:%s' % java_cm)
                exec_status,exec_output=subprocess.getstatusoutput(java_cm)
                if exec_status == 0:
                    
                    if os.path.exists(adjusting_all_file_name):
                        
                        tables=read_file(adjusting_all_file_name)
                        
                        all_data=adjust_tables(tables)
                        
                        o_data = output_data(all_data)
                                         
                        output_all_file_name=output_dir+file_name+'(adjusted)'+temp_file_suffix
                        with open(output_all_file_name,'wt',encoding='utf-8')  as tfile:
                            json.dump(o_data,tfile,ensure_ascii=False)
                            
                        # adjust_excel_file_name=output_dir+file_name+'(adjusted).xls'
                        # write_page_to_excel_by_sheet(adjust_excel_file_name,o_data)       
                        
                        os.remove(adjusting_all_file_name)
                else:
                    print('执行转换所有页出现异常。')
                    
                page_start = 1
                page_end = pages

                if need_check:                                   
                    
                    try:
                        check_page = find_param('page', 'top5')
                        print(check_page)
                        #need_check = True时，page参数有效。
                        #page的值几种表示方式：
                        #1.区间:2-5
                        #2.单页数字
                        #3.all
                        #4.top5 默认值
                        
                        if check_page.find('-') > 0:
                            sp_index = check_page.find('-')
                            page_start = int(check_page.split('-')[0])
                            page_end  = int(check_page.split('-')[1])
                            
                        else:
                            
                            if check_page != 'all':
                                
                                if check_page.find('top') == 0:
                                    page_start = 1
                                    page_end = int(check_page[len('top'):])
                                else:
                                    page_start = int(check_page)
                                    page_end = int(check_page)
                            else:
                                page_start = 1
                                page_end = pages
                                
                        if page_start > page_end or page_end > pages:
                            raise Exception('error value')
                            
                    except Exception:
                        print('错误的参数值.')
                        page_start = 1
                        page_end = 5
                    
                    
                    i=page_start
                    while i < page_end+1:
                        step=1
                    # for i in range(1,pages+1):
                        
                        print('正在解析文档：%s，第(%d)页。' %(os.path.split(file)[1], i))
                        lattice_table,exec_status=generation_page_file(adjusting_file_name,extract_mode, file_type ,i)
                        if exec_status!=0:
                            if exec_status == -100:
                                print(' Page number does not exist')
                                # pages=i-1
                                break
                                
                        page_table=adjust_tables(lattice_table)
                        
                        adjusted_page_file_name=output_dir+os.path.splitext(os.path.split(file)[1])[0]+'-page-'+repr(i)+'-(adjusted)'
                        excel_page_file_name = adjusted_page_file_name+'.xls'
                        json_page_file_name = adjusted_page_file_name+'.json'
                        
                        if os.path.exists(excel_page_file_name):
                            os.remove(excel_page_file_name)
                        if os.path.exists(json_page_file_name):
                            os.remove(json_page_file_name)
                        
                        #输出json
                        text_data = output_data(page_table)
                        if text_data:
                            with open(json_page_file_name, 'wt', encoding = 'gb18030') as fp:
                                json.dump(text_data, fp, ensure_ascii=False)
                        else:
                            print('该页未检测到表格数据。')
                        #写入excel
                        if page_table:
                            write_page_to_excel_by_sheet(excel_page_file_name, page_table)
                        
                        i+=step
                                
                print('文档%s已完成调整(%d-%d)页,共(%d)页。'%(os.path.split(file)[1], page_start, page_end, pages))
                        
            # print('Congratulations!all done!')
                   
        else:
                
            print('文件或目录:%s不存在.'%target)

            
if __name__ == '__main__':
    repair() 
