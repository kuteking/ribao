#打开一个文件取得开始的时间，将文件路径分解为目录路径文件，后缀，将时间转换为时间类型。
from tkinter import *
import os
import os.path

import tkinter.filedialog
import datetime
#from datetime import datetime,timedelta
from datetime import datetime,timedelta

def dedao_path():#获取头尾两个文件的路径，返回两个字符串
    file_dedao=tkinter.filedialog.askopenfilename()
    
    return file_dedao

def panduan(file_tou,file_wei):#对头尾文件判断是否存在
    try:
        os.path.isfile(file_tou)
        os.path.isfile(file_wei)
        return True
    except:
        return False







def file_path_time(file):
    
    file_path,wenjian_name=os.path.split(file)#file_path文件的路径
    wenjian_name,kuozhan_name=os.path.splitext(wenjian_name)#wenjian_name文件的名称。kuozhan_name 扩展名
    temp_time=''
    new_name=''
    for n in wenjian_name:
        if '0'<=n<='9':
            temp_time=temp_time+n
        else:
            new_name+=n
    if len(temp_time)>=7:
        wj_name_time=str_time(temp_time[0:4]+'-'+temp_time[4:6]+'-'+temp_time[6:])
    else:
        wj_name_time=None
    
    
    new_time=timestr(wj_name_time+timedelta(days=1))
    
    new_name=new_name+new_time+kuozhan_name
    new_file=file_path+new_name
    
    return new_file
    #return file_path,wenjian_name,wj_name_time,kuozhan_name#文件路径，文件名，文件名中的时间，文件扩展名/
#F:/', '潜江压气站压缩机管理日报表20190306', time.struct_time(tm_year=2019, tm_mon=3, tm_mday=6, /
#tm_hour=0, tm_min=0, tm_sec=0, tm_wday=2, tm_yday=65, tm_isdst=-1), '.xlsx')
def path_bijiao(file_tou,file_wei):#生成需要打开的文件的列表，返回列表
    if panduan(file_tou,file_wei):
        pass
    else:
        return False

    path_list=[]
    file_temp=file_tou
    
    while True:
        if os.path.samefile(os.path.abspath(file_temp),os.path.abspath(file_wei)):
            path_list.append([file_temp])
            return path_list
        else:
            path_list.append([file_temp])
            
            file_temp=file_path_time(file_temp)
            
            while True:
                if os.path.isfile(file_temp):
                    break
                else:
                    file_temp=file_path_time(file_temp)
                    
                
           


    
   
def str_time(data):#2006-03-02字符串转变为时间格式
    
    return datetime.strptime(data,"%Y-%m-%d")    
def timestr(datas):#时间转字符串
    return datas.strftime("%Y%m%d")

if __name__=='__main__':

    path_list=[]
    root=tkinter.Tk()
    print("tou")
    file_tou=dedao_path()
    print("wei")
    file_wei=dedao_path()
    
    path_list=path_bijiao(file_tou,file_wei)
    print(path_list)
    
    
    tkinter.mainloop()