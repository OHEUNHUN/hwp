# -*- coding: utf-8 -*-
"""
Created on Tue Oct 12 16:13:49 2021

@author: ECOCNA_dev
"""



import win32com.client as win32
import os
import pandas as pd
import dbupdate_update
import time



class chapter_3_Premission:

    def __init__(self):
        super().__init__()
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        
        
        self.samplepath = os.getcwd() + '\sample'
        self.filesample = "3_table_1.hwp"
        
        self.DbName = "hwp_3_1"
        
        self.filepath = os.getcwd() + '\Result\chapter3_허가배출기준'
        self.fileName = "chapter3_허가배출기준.hwp"



    def Get_property(self):
        '''
        초기 설정된 5개 항목에 대한 값을 불러 올 수 있다.
        
        '''
        return self.samplepath, self.filesample, self.DbName,self.filepath, self.fileName






    def Set_property(self,samplepath_in,filesample_in,DbName_in,filePath_in ,fileName_in):
        '''
        초기 설정을 변경할 수 있다.
        변경할 필요가 없다면 Enter를 쳐서 넘길 수 있다. 
        변경에 성공했다면 설정 성공리나는 말이 출력된다.
        실패하면 **없음으로 출력된다.
        
        '''
        
        
        try :
            if samplepath_in != "":
                self.samplepath = samplepath_in
                print("파일경로 설정 성공")
            else:
                self.samplepath = os.getcwd() + '\sample'
                print("설정된 경로 없음")
        
        
        
            if filesample_in != "":
                self.filesample = filesample_in
                print("파일샘플이름 설정 성공")
            else:
                self.filesample = "3_table_1.hwp"
                print("설정된 이름 없음")
        
        
        
            if DbName_in != "":
                self.DbName = DbName_in
                print("데이터베이스 이름 설정 성공")
            else:
                self.DbName_in = "hwp_3_1"
                print("설정된 이름 없음")
            
            
            
            if filePath_in != "":
                
                self.filepath = filePath_in
                print("파일이름 설정 성공")
                
            else:
                self.filepath = os.getcwd() + '\Result\chapter3_허가배출기준'
                print("설정된 이름 없음")
            
            
            
            if fileName_in != "":
                
                self.fileName = fileName_in
                print("파일이름 설정 성공")
                
            else:
                self.fileName = "chapter3_허가배출기준.hwp"
                print("설정된 이름 없음")
                
            print("새로운 경로 입력 완료")  
        
        except :
            
            print("경로 설정 오류")
    
    


    def hwpfile_open(self):
        
        
        try:
            self.hwp.XHwpWindows.Item(0).Visible = False
            self.hwp.RegisterModule("FilePathCheckDLL","FilePathcheckerModule")
        
        
            file_name = self.filesample
            sample_path = self.samplepath

            self.hwp.Open(os.path.join(sample_path,file_name))
            self.hwp.Run("MoveDocBegin") 


            self.hwp.Run("FileNew")
            self.hwp.XHwpWindows.Item(1).Visible = False
        except:
            print("해당 파일이 없습니다.")
            
            return "noflie"




    def Copy_Table_sample(self):
        
        '''
        원본에서 표를 복사해서 빈문서1에 붙여넣는다.
        
        '''
        self.hwp.XHwpDocuments.Item(0).SetActive_XHwpDocument()
        self.hwp.Run("SelectAll")
        self.hwp.Run("Copy")
        self.hwp.XHwpDocuments.Item(1).SetActive_XHwpDocument()
        self.hwp.Run("SelectAll")
        self.hwp.Run("Paste")




    def MakeTableCol(self,pollusion_list):

        num = 0

        table_list = ["배출구번호","주요배출시설","방지시설"]

        for i in range(len(pollusion_list)):
            num +=1
            self.hwp.MoveToField("배출구번호",select=False)
            self.hwp.Run("TableInsertLowerRow")

            if num>1:
                for j in table_list:
                    self.hwp.MoveToField(j,select=True)
                    self.hwp.Run("TableLowerCell")
                    self.hwp.Run("TableCellBlockExtend")
                    self.hwp.Run("TableLowerCell")
                    self.hwp.Run("TableMergeCell")




    def inserttext(self,text):
        self.hwp.Run("SelectAll")
        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
        self.hwp.HParameterSet.HInsertText.Text = text
        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)




    def InsertText_stack(self,Select_data):

        self.hwp.MoveToField("배출구번호",select=False)
        self.hwp.Run("TableLowerCell")
        self.inserttext(Select_data[0])

        self.hwp.MoveToField("주요배출시설",select=False)
        self.hwp.Run("TableLowerCell")
        self.inserttext(Select_data[1])

        self.hwp.MoveToField("방지시설",select=False)
        self.hwp.Run("TableLowerCell")
        self.inserttext(Select_data[2])



    def InsertText_chem(self, Select_data):
        table_list =["오염물질","최대배출기준","허가배출기준","예상배출농도","단위","관리항목","최대배출기준근거"]
        
        for i,data_list in enumerate(Select_data):

            for j, data in enumerate(data_list):

                self.hwp.MoveToField(table_list[j],select=False)

                for k in range(i+1):
                    self.hwp.Run("TableLowerCell")

                self.inserttext(data)







    def DeleteFristCell(self):
        self.hwp.MoveToField("배출구번호",select=True)
        self.hwp.Run("TableCellBlockExtend")
        self.hwp.Run("TableColEnd")
        self.hwp.SetMessageBoxMode(0x00002000)
        self.hwp.Run("TableDeleteCell")




    def InsertCell(self,data_list):

        stack = data_list.drop_duplicates([0])
        stack = stack.sort_index(ascending=False)
        stack = stack[[0,1,2]].reset_index(drop=True)

        for i in stack[0]:

            item_chem = data_list.loc[data_list[0]==i][[3,4,5,6,7,8,9]]
            item_chem = item_chem.reset_index(drop=True)
            item_chem_list_front = item_chem[3].values.tolist()
            item_chem_list_back = item_chem.values.tolist()

            item_stack = data_list.loc[data_list[0]==i][[0,1,2]]
            item_stack = item_stack.reset_index(drop=True)
            item_stack_list = item_stack.values.tolist()[0]

            self.MakeTableCol(item_chem_list_front)
            self.InsertText_stack(item_stack_list)
            self.InsertText_chem(item_chem_list_back)


            print(i)
            print(item_chem)

        self.DeleteFristCell()




    def Main(self,host,port,database,user,password):
        
        dbname = self.DbName
        

        db = dbupdate_update.DBCONN()
        db.SetDataBase(host,port,database,user,password)
        db.DB_CONN()
        
        hwp_table = db.query(dbname, "*")
        hwp_table = pd.DataFrame(hwp_table)
        hwp_table.fillna(" ",inplace=True)

        self.hwpfile_open()
        self.Copy_Table_sample()
        self.InsertCell(hwp_table)

        path = os.path.abspath(self.filepath)
        name = self.fileName


        self.hwp.XHwpWindows.Item(1).Visible = True
        time.sleep(0.2)

        self.hwp.SaveAs(os.path.join(path,name))

        time.sleep(0.2)
        self.hwp.Run("FileClose")
        time.sleep(0.2)
        
        self.hwp.XHwpWindows.Item(0).Visible = True
        time.sleep(0.2)
        self.hwp.Run("FileClose")
        time.sleep(0.2)

        self.hwp.Quit()


