# -*- coding: utf-8 -*-
"""
Created on Mon May 21 15:30:25 2018

@author: Dell
"""
import ctypes    
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")  
from tkinter import *
import tkinter.filedialog
import tkinter
import hashlib
import time
import xlrd
LOG_LINE_NUM = 0




class MY_GUI():
    def __init__(self,init_window_name):
        self.init_window_name = init_window_name
        self.all_table_index_list=[]
        self.path=""    
        self.all_materi_set=set()  #所有的材料类别
        self.all_finacial_category=set() #所有的财务科目
    #设置窗口
    def set_init_window(self):
        self.init_window_name.title("财务计算小程序-周凝倩")           #窗口名
        self.init_window_name.geometry('320x160+10+10')                         #290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.init_window_name.geometry('1400x800')
        self.init_window_name["bg"] = "pink"                                    #窗口背景色，其他背景色见：blog.csdn.net/chl0000/article/details/7657887
        self.init_window_name.attributes("-alpha",0.9)                          #虚化，值越小虚化程度越高
        self.init_window_name.iconbitmap('ico.ico')     #指定界面图标
        #标签
        self.init_data_label = Label(self.init_window_name, text="输入的excel表格的位置")
        self.init_data_label.grid(row=0, column=1)
        self.result_data_label = Label(self.init_window_name, text="输出结果")
        self.result_data_label.grid(row=0, column=8)
        self.log_label = Label(self.init_window_name, text="计算过程及操作日志")
        self.log_label.grid(row=10, column=1)
        #文本框
        self.init_data_Text = Text(self.init_window_name, width=50, height=0)  #原始数据录入框
        self.init_data_Text.grid(row=1, column=1, rowspan=1, columnspan=2)
        self.result_data_Text = Text(self.init_window_name, width=70, height=49)  #处理结果展示
        self.result_data_Text.grid(row=1, column=8, rowspan=16, columnspan=12)
        self.log_data_Text = Text(self.init_window_name, width=90, height=35)  # 日志框
        self.log_data_Text.grid(row=13, column=0, rowspan=10,columnspan=6)
        
        
        
        #按钮
        self.str_trans_to_md5_button = Button(self.init_window_name, text="清除文件", bg="red", width=20,command=self.inputclear)  # 调用内部方法  加()为直接调用
        self.str_trans_to_md5_button.grid(row=1, column=4)
        self.openfilebutton = Button(self.init_window_name, text="打开文件", bg="lightblue", width=10,command=self.open_file)
        self.openfilebutton.grid(row=1, column=3)
        
        self.calculate_all = Button(self.init_window_name, text="计算所有总和(并初始化计算资源)", bg="lightblue", width=25,command=self.calculate_all_sum)  # 调用内部方法  加()为直接调用
        self.calculate_all.grid(row=5, column=2)
        self.calculate_all = Button(self.init_window_name, text="清除计算结果", bg="yellow", width=15,command=self.clear_output)  # 调用内部方法  加()为直接调用
        self.calculate_all.grid(row=5, column=5)
        
        self.calculate_all = Button(self.init_window_name, text="查看科目名称", bg="lightblue", width=15,command=self.check_all_finacial)  # 调用内部方法  加()为直接调用
        self.calculate_all.grid(row=5, column=3)
        self.calculate_all = Button(self.init_window_name, text="查看材料款名称", bg="lightblue", width=15,command=self.check_all_material)  # 调用内部方法  加()为直接调用
        self.calculate_all.grid(row=6, column=3)
        self.calculate_all = Button(self.init_window_name, text="计算所财务科目明细", bg="lightblue", width=15,command=self.get_catrgory_all)  # 调用内部方法  加()为直接调用
        self.calculate_all.grid(row=5, column=4)
        self.calculate_all = Button(self.init_window_name, text="计算(材料款)明细", bg="lightblue", width=15,command=self.cal_material_minxi)  # 调用内部方法  加()为直接调用
        self.calculate_all.grid(row=6, column=4)
        


        
        #滚动条
        self.result_data_scrollbar_y = Scrollbar(self.init_window_name)    #创建纵向滚动条
        self.result_data_scrollbar_y.config(command=self.result_data_Text.yview)  #将创建的滚动条通过command参数绑定到需要拖动的Text上
        self.result_data_Text.config(yscrollcommand=self.result_data_scrollbar_y.set)
        self.result_data_scrollbar_y.grid(row=1, column=25, rowspan=20, sticky='NS')
        
        self.log_data_scrollbar_y = Scrollbar(self.init_window_name)    #创建纵向滚动条
        self.log_data_scrollbar_y.config(command=self.log_data_Text.yview)  #将创建的滚动条通过command参数绑定到需要拖动的Text上
        self.log_data_Text.config(yscrollcommand=self.log_data_scrollbar_y.set)
        self.log_data_scrollbar_y.grid(row=13, column=5, rowspan=15, sticky='NS')
        
        
        
        
    def open_file(self):
        self.path=tkinter.filedialog.askopenfilename()
        print(self.path)
        self.init_data_Text.insert(END, self.path)
        self.write_log_to_Text("信息:加载文件%s成功!"%self.path)
    def inputclear(self):
            
        self.init_data_Text.delete(1.0,END)
        self.write_log_to_Text("信息:清除文件%s成功"%self.path)
        self.path=" " 
    def clear_output(self):
        
        self.result_data_Text.delete(1.0,END)
        self.write_log_to_Text("信息:清除结果信息成功功")
        
    def get_sheet_start_and_end(self,sh,keyword_rows,keword_cols):
        nrows = sh.nrows  # 获取行数
        ncols = sh.ncols  # 获取列数
        real_rows=-1  #真实的行
        real_cols=-1  #真实的列
        for i in range(0,nrows):          #循环第一列的每一行，出现了“合计”表示最后一行
            cell_name=sh.cell_value(i,0)
            if cell_name==keyword_rows:
                # print i
                real_rows=i
                break
        for j in range(0,ncols):   #循环第三行的每一列，表示出现了“备注”表示最后一列了
            cell_name=sh.cell_value(3,j)
            if cell_name==keword_cols:
                # print j
                real_cols=j
                break
        return real_rows,real_cols
    
    def calculate_all_sum(self):
        self.all_table_index_list=[] #必须初始化，防止重复计算利用
        filename = self.path
        print(self.path)
        bk = xlrd.open_workbook(filename)
        n = bk.nsheets
        shrange = range(bk.nsheets)
        reault=0.0
        self.result_data_Text.insert(END,"\n")
        print (u"表格名称",u"合计",u"金额")
        self.write_log_to_Text("开始解析每个表的数据，计算每个表的总金额")
        strs_titte = str("表格名称") +"          " + str("合计") +"          "+str("金额") + "\n"      #换行
        self.write_log_to_Text(strs_titte)
        for sheet in shrange:
            #sh为获取到的sheet的内容
            sh=bk.sheet_by_index(sheet)
            keyword_rows=u"合计"
            keword_cols=u"备注"
            real_rows,real_cols=self.get_sheet_start_and_end(sh,keyword_rows, keword_cols)
            # print  (real_rows, real_cols)
            self.all_table_index_list.append([sheet,real_rows,real_cols])
            #  第三行开始是抬头名称
            cell_name=sh.cell_value(real_rows, 0)
            cell_value = sh.cell_value(real_rows, real_cols-1)
            print (sh.name,cell_name,cell_value)
            insert_format="%s    %s    %.2f"%(sh.name,cell_name,cell_value)
            self.write_log_to_Text(insert_format)
            reault+=cell_value
        print( u"所有的合计",reault)
        self.write_log_to_Text(" 所有表计算的总金额为：%s"%reault)
        self.result_data_Text.insert(END,"\n所有的表的合计金额为：")
        self.result_data_Text.insert(END,reault)
        self.result_data_Text.insert(END,"\n")
        self.result_data_Text.insert(END,"*********")
        self.result_data_Text.insert(END,"\n")
        return reault
    
    def check_all_finacial(self):
        self.write_log_to_Text(" ")
        self.write_log_to_Text("查看所有财务类别信息")
        self.result_data_Text.insert(END,"一共有以下的财务类别：")
        self.result_data_Text.insert(END,"\n")
        bk = xlrd.open_workbook(self.path)
        for index in self.all_table_index_list:
            sh = bk.sheet_by_index(index[0])
            ncols = index[2]
            for j in range(1, ncols):
                temp = sh.cell_value(3, j)
                if temp not in [u"小计", u"合计","材料类别"]:
                    self.all_finacial_category.add(temp)
        
        for leibie in self.all_finacial_category:
            self.write_log_to_Text("%s"%leibie)
            self.result_data_Text.insert(END,leibie)
            self.result_data_Text.insert(END,"\n")
        self.result_data_Text.insert(END,"*********")
        self.result_data_Text.insert(END,"\n")
        return self.all_finacial_category
    def check_all_material(self):
        self.write_log_to_Text(" ")
        self.write_log_to_Text("查看所有材料类别信息")
        self.result_data_Text.insert(END,"一共有以下的材料款类别：")
        self.result_data_Text.insert(END,"\n")
        bk = xlrd.open_workbook(self.path)
        for index in self.all_table_index_list:
            sh = bk.sheet_by_index(index[0])
            nrows=index[1]
            ncols = index[2]
            for i in range(3,nrows):
                temp=sh.cell_value(i, 1)
                if type(temp)==float:
                    self.all_materi_set.add(temp)
        for cailiao in self.all_materi_set:
            if cailiao!=u"材料类别":
                self.write_log_to_Text("%s"%cailiao)
                self.result_data_Text.insert(END,cailiao)
                self.result_data_Text.insert(END,"\n")
        self.result_data_Text.insert(END,"*********")
        self.result_data_Text.insert(END,"\n")
        return self.all_materi_set
    
    def cal_material_minxi(self):
        self.write_log_to_Text(" ")
        self.write_log_to_Text("查看所有材料类别的总合计及明细")
        self.result_data_Text.insert(END,"查看所有材料类别的总合计及明细:")
        self.result_data_Text.insert(END,"\n")
        bk = xlrd.open_workbook(self.path)
        matreial_category_count={}
        #求材料款个大类别合计
        if len(self.all_materi_set)<1:
            self.check_all_material()
        for id in self.all_materi_set:
            matreial_category_count[id]=0.0
        #每一个类别的总量
        cailiao_sum=0.0
        x=0.0
        for index in self.all_table_index_list:      
            sh = bk.sheet_by_index(index[0])
            nrows=index[1]
            ncols = index[2]
            key = sh.cell_value(nrows-1, 1)
            for indey in range(1,ncols):
                name=sh.cell_value(3, indey)     
                if name==u"材料款":       
                    value = sh.cell_value(nrows, indey)       
                    matreial_category_count[key]+=value
                    cailiao_sum+=value

        print (u"材料款各类别合计",":",cailiao_sum  )           
        self.result_data_Text.insert(END,cailiao_sum)
        self.result_data_Text.insert(END,"\n")
        self.write_log_to_Text("总和")
        self.write_log_to_Text(cailiao_sum)
        print (u"材料类别明细\n")
        for id in matreial_category_count:
            print (u"材料类别",id,":",matreial_category_count[id])
            strs_materis="%s    %.2f"%(id,matreial_category_count[id])
            self.result_data_Text.insert(END,strs_materis)
            self.result_data_Text.insert(END,"\n")
            self.write_log_to_Text(strs_materis)
        
        
    def get_catrgory_all(self):
        bk = xlrd.open_workbook(self.path)
        key_word=self.all_finacial_category
        if len(self.all_finacial_category)<1:
            key_word=self.check_all_finacial()
        finacial_dict_couont={}
        for id in key_word:
            finacial_dict_couont[id]=0.0   
        for index in self.all_table_index_list:
            # print index[0],"*"*20
            sh = bk.sheet_by_index(index[0])
            nrows=index[1]
            ncols = index[2]
            for j in range(2,ncols-1):
                key=sh.cell_value(3, j)
                value=sh.cell_value(nrows,j)
                if key in key_word:
                    finacial_dict_couont[key]+=value
        # print finacial_dict_couont
        print (u" 各个科目明细")
        self.write_log_to_Text("%s"%"各个科目明细")
        self.result_data_Text.insert(END,"各个科目明细")
        self.result_data_Text.insert(END,"\n")
        for name in finacial_dict_couont.keys():
            print (name,finacial_dict_couont[name])
            minxikemu="%10s    %.2f"%(name,finacial_dict_couont[name])
            self.write_log_to_Text("%s"%minxikemu)
            self.result_data_Text.insert(END,minxikemu)
            self.result_data_Text.insert(END,"\n")
        return finacial_dict_couont
    #功能函数
    def str_trans_to_md5(self):
        src = self.init_data_Text.get(1.0,END).strip().replace("\n","").encode()
        #print("src =",src)
        if src:
            try:
                myMd5 = hashlib.md5()
                myMd5.update(src)
                myMd5_Digest = myMd5.hexdigest()
                #print(myMd5_Digest)
                #输出到界面
                self.result_data_Text.delete(1.0,END)
                self.result_data_Text.insert(1.0,myMd5_Digest)
                self.write_log_to_Text("INFO:str_trans_to_md5 success")
            except:
                self.result_data_Text.delete(1.0,END)
                self.result_data_Text.insert(1.0,"字符串转MD5失败")
        else:
            self.write_log_to_Text("ERROR:str_trans_to_md5 failed")


    #获取当前时间
    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        return current_time


    #日志动态打印
    def write_log_to_Text(self,logmsg):
        global LOG_LINE_NUM
        current_time = self.get_current_time()
        logmsg_in = str(current_time) +" " + str(logmsg) + "\n"      #换行
        self.log_data_Text.insert(END, logmsg_in)


def gui_start():
    init_window = Tk()              #实例化出一个父窗口
    ZMJ_PORTAL = MY_GUI(init_window)
    # 设置根窗口默认属性
    ZMJ_PORTAL.set_init_window()

    init_window.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


gui_start()