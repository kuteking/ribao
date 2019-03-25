#读写excel（xlrd、xlwt）
import xlrd
import xlwt
import os
import os.path
import tkinter.filedialog

def excel_rd(file,liebiao):#要打开的ecxel的文件 file  rd_list要读取数据的的表格坐标列表
    temp_list=[]
    woexcel=xlrd.open_workbook(file)
    excel_data=woexcel.sheets()[0]
    lie=excel_data.ncols
    hang=excel_data.nrows
    for n in liebiao:
        hh=n
        temp_list.append((excel_data.cell_value(hh,7)+'\n'))
    return temp_list

def data_wr(data_list):
    
    op=open('f:\\testexcel.txt',"w")
    for n in data_list:
        #op.write(",".join('%s' %id for id in n)+"\n")
        op.write(str(n))
        
    op.close()

def sj_cl(data_list):#洗出时间
    time_list = []
    bj = False
    xiaoshi=''
    fenzhong=''
    for nn in data_list:
        for n in nn:
            if n == '小':
                bj = True
            if '0'<=n<='9' and bj == False:
                xiaoshi=xiaoshi+n
            if '0'<=n<='9' and bj == True:
                fenzhong=fenzhong+n
        #if fenzhong[0] =='0':
            #fenzhong=fenzhong[1]
        
        time_list.append(xiaoshi+'.'+fenzhong)
        xiaoshi=''
        fenzhong=''
        bj = False
    
    return time_list#返回的是分钟数




def main():
    data_list=[]
    excel_biao=[3,4,5,6,24,25,26,46,47,48,49]
    file_temp=tkinter.filedialog.askopenfilename()
   
    data_list=excel_rd(file_temp,excel_biao)
    data_wr(data_list)
    return sj_cl(data_list)
    #print(sj_cl(data_list))


