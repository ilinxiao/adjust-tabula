import Tools
from Tools import *
import  xdrlib ,sys
import xlrd,xlwt
from functools import reduce
import re
import json
import subprocess,os
import shutil
import hashlib
from  hashlib import md5

table_header={
                                'feixingqingbaoqu':['名称、地名代码','水平范围','备注'],
                                'quyuguanzhiqu':['名称、水平范围、垂直范围', '提供服务的单位', '呼号', '工作频率(*表示备用)', '服务时间', '备注'],
                                'zhongduanguanzhiquhejinjinguanzhiqu':['名称、范围、高度','提供服务的单位','话呼','工作频率、时间','备注'],
                                'Gxiliehangxian':['航路、航线代号、导航点名称、坐标','磁航迹距离(千米/海里)','最低飞行高度(米)','宽度(千米)','巡航高度层方向','管制单位'],
                                'Jxiliehangxian': ['航路、航线代号、导航点名称、坐标','磁航迹距离(千米/海里)','最低飞行高度(米)','宽度(千米)','巡航高度层方向','管制单位'],
                                'baogaodianbiao':['名称代码','坐标','航路'],
                                'zhifubao':['基本信息']}
#一页一个sheet
def write_page_to_excel_by_sheet(filename,tables):

    pages=len(tables)
    book=xlwt.Workbook()
    
    for i in range(0,len(tables)):
        
        table=tables[i]
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
    
    # write_excel('mark_data_preview.xls',mark_data)
    #4.移动
    new_table=[]
    
    # print('new_table_header:%s' % str(new_table))
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
                                    Tools.output_file('debug_the_same_cell.txt','ocell text: %s' % ocell['text']+'\ncell text:%s' % cell['text']+'\n')
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
  
global print_once
print_once=True
def merge_cells(cells):
    global print_once
    #去重複元素的高深莫测写法~~
    func=lambda x,y:x if y in x else x+[y]
    left_list=reduce(func,[[],]+[obj['left'] for obj in cells])
    # print(left_list)

    temp_cells=sorted(cells,key=lambda obj:obj['top'])
    # row_max_height=max([obj['height'] for obj in temp_cells])
    # if print_once:
        # print('-'*25+'开始处理：'+'-'*25)
        # print(temp_cells)
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
            # if col_merge_height!=row_max_height:
                # print('合并高度有问题哟~')
                # print('col_merge_height:%s/max_row_height:%s'%(str(col_merge_height)))
            # else:
    # if print_once:
        # print('-'*25+'合并列之后：'+'-'*25)
        # print(temp_cells)
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
    # print(row_min_obj)
    # if print_once:
        # print('-'*25+'合并结束：'+'-'*25)
        # print(temp_cells)
    # if len(temp_cells)!=1:
        # print('合并结果：,len(result_cells):%d' % len(temp_cells))
    # print_once=False   
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
    
def set_empty_cell(cell): 
    cell['left']=0.0
    cell['width']=0.0
    cell['text']=''
    cell['height']=0.0
    cell['top']=0.0
# def move_cell(row,start,end):
    # if start!=end:
def read_json_data(filename):
    tables=[]
    with open(filename,'r',encoding='gb18030') as file:
        tables=json.load(file)
    return tables      
    
def is_empty_table(table):
    
    if 'width' in table.keys() and 'height' in table.keys():
        if table['width'] == 0 or table['height'] ==0:
            return True
    return False
    
def filter_empty_table(tables):
    
    temp_tables=tables[:]
    empty_data=[]
    for i in range(len(temp_tables)-1,-1,-1):
        if is_empty_table(temp_tables[i]):
            empty_data.append(temp_tables[i])
            temp_tables.pop(i)
            
    return temp_tables,empty_data
  
def has_data(table):
    if type(table) == type({}):
        if 'data' in table.keys():
            return True
    return False   

#提取data    
def filter_data(tables):
    
    temp_tables=tables[:]
    data_data=[]
    no_data_data=[]
    for i in range(len(temp_tables)-1,-1,-1):
        if has_data(temp_tables[i]):    
            data_data.append(temp_tables[i]['data'])
        else:
            no_data_data.append(temp_tables[i])
            temp_tables.pop[i]
                
    return data_data,no_data_data
    

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
    
def get_pages(tables):
    header_map={}
    for table in tables:
        data=table['data']

        for i in range(0,len(data)):
            count=1
            obj={}
            index_list=[]
            row=data[i]
            # if i==0:
                # print([obj['text'] for obj in row])
                # print(get_row_hash(row))
            row_text=get_row_text(row)
            for key,header in table_header.items():
                # print('key:%s,value:%s' % (key,header))
                if row_text == ''.join(header):
                    if key in header_map.keys():
                        count=header_map[key]
                        count+=1
                    header_map[key]=count
                                   
    # reduce(lambda x,y)
    return header_map
    # print([obj['row'] for obj in row_hash.values() if obj['count'] == max([c['count'] for c in row_hash.values()]) or obj['index_list'].count(0)/len(obj['index_list'])>0.8])
def is_header(row):
    row_text=get_row_text(row)
    for key,header in table_header.items():
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

def get_page_file_name(file_name,extract_mode,page):
    return os.path.splitext(file_name)[0]+'-'+repr(page)+'-'+extract_mode+os.path.splitext(file_name)[1]
    
def generation_page_file(tabula_path,target_file_name,extract_mode,file_type,page,file_name):
    temp_file_name=get_page_file_name(file_name,extract_mode,page)
    # print('第%d页：%s' % (page,temp_file_name))
    java_cm='java -jar %s %s -%s -f %s -p %d -o %s' %(tabula_path,target_file_name,extract_mode,file_type,page,temp_file_name)
    exec_status,exec_output=subprocess.getstatusoutput(java_cm)
    table=[]
    if exec_status==0:
        if os.path.exists(temp_file_name):
            table=read_file(temp_file_name)
            # print('generation_page_file/read_file:%s' % table)
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
    # return ''.join([obj['text'] for obj in row])
# sort_exception=read_json_data('../temp/da5fc44fd3bf9cfc6e47e1296bf1375a-5-r.json')
# text_data=output_data(sort_exception)[0]
# print(text_data)
# write_excel('../temp/shijiazhuang.xls',)
# no_empty_data,empty_data=filter_empty_table(read_json_data('../../PART-p-52-r.json'))
# adjusted_data=adjust_tables(no_empty_data)
# write_page_to_excel_by_sheet('PART-p-52-r.xls',adjusted_data)
# temp_json_data=read_json_data('../temp/ee689539863f0518dbd638adc26411cf-2-r.json')
# temp_json_data=read_json_data('sort_exception.json')
# no_empty_data,empty_data=filter_empty_table(temp_json_data)

# text_data=output_data(no_empty_data)
# print(text_data)
# with open('sort_exception(preview).json','wt',encoding='utf-8')  as tfile:
     # json.dump(text_data,tfile,ensure_ascii=False)
# no_empty_data,empty_data=filter_empty_table(read_json_data('../PART-all-r.json'))
# Tools.output_file('PART-all-r_adjust.txt',output_data(adjust_tables(no_empty_data)))
# Tools.output_file('Part2-all-r_adjust.txt',output_data(adjust_tables(no_empty_data)))
# data_data,no_data_data=filter_data(no_empty_data)       
# no_empty_data,empty_data=filter_empty_table(read_json_data('../313PART2-all-r.json'))
# print(get_pages(no_empty_data))
# Tools.output_file('313PART2-all-r_text_data.txt',output_data(adjust_tables(no_empty_data)))
# write_excel('adjust_table_example_data_test1.xls',finally_data)
# adjust(table2)
# write_excel('adjust_table_example_data.xls',table2)
# print(sys.argv)

def repair():

    if len(sys.argv)>1:
        target=sys.argv[1]
        print(target)
        
        if os.path.exists(target):
        
            pdf_files=[]
            tabula_path='../tabula/tabula-1.0.1-jar-with-dependencies.jar'
            print('target is exists.')
            if os.path.isfile(target):
                # print('a file.')
                if os.path.splitext(target)[1] == '.pdf':
                    pdf_files.append(target)
                    
            elif os.path.isdir(target):
                # print('is dir.')

                for file in os.listdir(target):
                    file_path=os.path.join(target,file)
                    # print(file)
                    # print(os.path.splitext(file_path)[1] == 'pdf')
                    # print('读取到的pdf文件有：')
                    if os.path.isfile(file_path) and os.path.splitext(file_path)[1] == '.pdf':
                        pdf_files.append(file_path)
                        
            for file in pdf_files:
                
                # print('splitext(file):%s' % str(os.path.splitext(file)))
                # print('split(file):%s' % str(os.path.split(file)))
                temp_file_path,file_suffix=os.path.splitext(file)
                # print('temp_file_path:%s' % temp_file_path)
                # print('temp_file_name:%s' %temp_file_name)
                file_name_hash=hashlib.md5(temp_file_path.encode('utf-8')).hexdigest()
                temp_dir='../temp/'
                if not os.path.exists(temp_dir):
                    os.mkdir(temp_dir)
                
                file_name=os.path.join(temp_dir, file_name_hash+file_suffix)
                # print('file_name_hash:%s' % file_name)
                if os.path.exists(file_name):
                    os.remove(file_name)
                shutil.copy2(file,file_name)
                
                # if os.path.exists(file_name):
                    # print('%s复制完成。' % file_name)
                    
                file_type='JSON'
                temp_file_suffix=''
                if file_type == 'JSON':
                    temp_file_suffix='.json'
                elif file_type == 'CSV':
                    temp_file_suffix='.csv'
                 
                #先解析出全部数据 获取页数信息
                all_file_name=temp_dir+file_name_hash+'all'+temp_file_suffix
                java_cm=u'java -jar %s %s -r -f %s -p %s -o %s' %(tabula_path,file_name,file_type,'all',all_file_name)
                
                exec_status,exec_output=subprocess.getstatusoutput(java_cm)
                if exec_status == 0:
                    
                    if os.path.exists(all_file_name):
                        
                        tables=read_file(all_file_name)
                        page_info=get_pages(tables)
                        pages=0
                        for num in page_info.values():
                            pages+=num
                     # json_file=temp_dir+file_name_hash+temp_file_suffix
                    # print('json_file:%s' % json_file)
                        all_data=[]
                                            
                        output_dir='../output/'
                        if not os.path.exists(output_dir):
                            os.mkdir(output_dir)
                        print('正在解析文档：%s' %(os.path.split(file)[1]))
                        i=1
                        while i<pages+1:
                            step=1
                        # for i in range(1,pages+1):
                            
                            lattice_table,exec_status=generation_page_file(tabula_path,file_name,'r',file_type,i,temp_dir+file_name_hash+temp_file_suffix)
                            if exec_status!=0:
                                if exec_status == -100:
                                    #get_pages缺陷是无法准确的知道一页有多少个表格
                                    print(' Page number does not exist')
                                    pages=i-1
                                    break
                                    
                            stream_table=generation_page_file(tabula_path,file_name,'t',file_type,i,temp_dir+file_name_hash+temp_file_suffix)
                            
                            # print(lattice_table)
                            # if has_header(lattice_table):
                            new_table=adjust_tables(lattice_table)
                            # if len(new_table)>1:
                               # pages-len(new_table)+1
                                # print('同一页中有%d个表格。' % len(new_table))
                            
                            # print('new_table:%s' % new_table)
                            corr_data=[]
                            for xx in range(0,len(new_table)):
                                if xx >0:
                                     pages-=1
                                else:
                                    print('第%d页' %i)
                                corr_data.append(correction_data(new_table[xx],stream_table))
                            # print('length of corr_data：%d' % len(corr_data))

                            # print('check new_table length:%d' % len(new_table))
                            #输出每页调整后的数据
                            # page_adjusted_file_name=output_dir+os.path.splitext(os.path.split(file)[1])[0]+'-page-'+repr(i)+'-(adjusted)'+temp_file_suffix
                            # with open(page_adjusted_file_name,'wt',encoding='utf-8')  as tfile:
                                # json.dump(corr_data,tfile,ensure_ascii=False)
                            # print('调整替换后的数据：%s' % corr_data)
                            page_adjusted_file_name=output_dir+os.path.splitext(os.path.split(file)[1])[0]+'-page-'+repr(i)+'-(adjusted).xls'
                            
                            if os.path.exists(page_adjusted_file_name):
                                # print('文件：%s已被删除.' % page_adjusted_file_name)
                                os.remove(page_adjusted_file_name)
                                
                            write_page_to_excel_by_sheet(page_adjusted_file_name,corr_data)
                            # else:
                                # print('no header found.')
                            text_table=output_data(corr_data)
                            
                            all_data.extend(text_table)
                            i+=step
                            
                        adjust_file_name=output_dir+os.path.splitext(os.path.split(file)[1])[0]+'(adjusted)'+temp_file_suffix
                        with open(adjust_file_name,'wt',encoding='utf-8')  as tfile:
                            json.dump(all_data,tfile,ensure_ascii=False)
                            
                        adjust_excel_file_name=output_dir+os.path.splitext(os.path.split(file)[1])[0]+'(adjusted).xls'
                        write_page_to_excel_by_sheet(adjust_excel_file_name,all_data)   
                         
                os.remove(all_file_name)
                print('文档%s共%d页数据提取完成。'%(file,pages))
                        
            print('Congratulations!all done!')
                   
        else:
                
            print('文件或目录:%s不存在.'%target)

repair() 