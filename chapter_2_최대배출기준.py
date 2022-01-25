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

class chapter_2_Max:

    def __init__(self):
        super().__init__()

        
        self.samplepath = os.getcwd() + '\sample'
        self.filesample = "2_table_1.hwp"
        self.DbName = "hwp_2_2_1"
        self.filepath = os.getcwd() + '\Result\chapter2_최대배출기준'
        self.fileName = "chapter2_최대배출기준.hwp"





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
                self.samplepath = os.getcwd()+ '\sample'
                print("설정된 경로 없음")
        
        
        
            if filesample_in != "":
                self.filesample = filesample_in
                print("파일샘플이름 설정 성공")
            else:
                self.filesample = "2_table_1.hwp"
                print("설정된 이름 없음")
        
        
        
            if DbName_in != "":
                self.DbName = DbName_in
                print("데이터베이스 이름 설정 성공")
            else:
                self.DbName = "hwp_2_2_1"
                print("설정된 이름 없음")
            
            
            
            if filePath_in != "":
                
                self.filepath = filePath_in
                print("파일이름 설정 성공")
                
            else:
                self.filepath = os.getcwd() + '\Result\chapter2_최대배출기준'
                print("설정된 이름 없음")
            
            
            if fileName_in != "":
                
                self.fileName = fileName_in
                print("파일이름 설정 성공")
                
            else:
                self.fileName = "chapter2_최대배출기준.hwp"
                print("설정된 이름 없음")
                
            print("새로운 경로 입력 완료")  
        
        except :
            
            print("경로 설정 오류")
        



    def hwpfile_open(self):
        '''
        샘플파일을 연다. 
        파일 경로와 이름은 위에서 초기 설정서에서 관리하며, 변경이 필요한 경우 Set_property 메소드를 사용한다.
        
        '''
        try :
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

        table_list = ["배출구번호","배출구종구분","x좌표","y좌표","표고","용마루높이","굴뚝높이","내경","유속",
        "배가스온도","배가스유량","배출오염물질개수"]

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
        table_list = ["배출구번호","배출구종구분","x좌표","y좌표","표고","용마루높이","굴뚝높이","내경","유속",
        "배가스온도","배가스유량","배출오염물질개수"]


        for i,data in enumerate(Select_data):
            self.hwp.MoveToField(table_list[i],select=False)
            self.hwp.Run("TableLowerCell")
            self.inserttext(data)
    
    def InsertText_chem(self, Select_data):
        table_list =["오염물질","배출농도","배출량","배출영향분석"]
        
        for i,data in enumerate(Select_data):
            self.hwp.MoveToField(table_list[0],select=False)

            for j in range(i+1):
                self.hwp.Run("TableLowerCell")
            
            self.inserttext(data[0])
            self.hwp.Run("TableRightCellAppend")
            self.inserttext(data[1])
            self.hwp.Run("TableRightCellAppend")
            self.inserttext(data[2])
            self.hwp.Run("TableRightCellAppend")
            self.inserttext(data[3])
            


    def InsertCell(self,data_list):
        
        stack = data_list.drop_duplicates([0])
        stack = stack.sort_index(ascending=False)
        stack = stack[0].reset_index(drop=True)


        for i in stack:
            item = data_list.loc[data_list[0] == i]

            #["배출구번호","배출구종구분","x좌표","y좌표","표고","용마루높이","굴뚝높이","내경","유속","배가스온도","배가스유량","배출오염물질개수"]
            #위에 데이터에 들어갈 정보로 만든 데이터프레임
            item1 = item.loc[: , 0:11].drop_duplicates()
            #화학종에 대한 정보로 만든 데이터프레임
            item2 = item.loc[:, 12:15]

            #위에 항목을 리스트로 변경
            item1_list = item1.values.tolist()[0]

            #화학종만 분리해서 만든 리스트(표를 만들기 위함)
            item2_list = item2[12].tolist()

            #화학종에 들어갈 모든 데이터
            item3_list = item2.values.tolist()


            self.MakeTableCol(item2_list)
            self.InsertText_stack(item1_list)
            self.InsertText_chem(item3_list)

            print(item)

    


    def DeleteFristCell(self):
        self.hwp.MoveToField("배출구번호",select=True)
        self.hwp.Run("TableCellBlockExtend")
        self.hwp.Run("TableColEnd")
        self.hwp.SetMessageBoxMode(0x00002000)
        self.hwp.Run("TableDeleteCell")





    def Main(self,host,port,database,user,password):


        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        dbname = self.DbName

        db = dbupdate_update.DBCONN()
        db.SetDataBase(host,port,database,user,password)
        db.DB_CONN()

        hwp_table = db.query (dbname,"*")
        hwp_table = pd.DataFrame(hwp_table)
        hwp_table = hwp_table.fillna(" ")


        self.hwpfile_open()
        self.Copy_Table_sample()
        self.InsertCell(hwp_table)
        self.DeleteFristCell()

        path = os.path.abspath(self.filepath)
        name = self.fileName

        self.hwp.XHwpWindows.Item(1).Visible = True
        time.sleep(0.2)
        self.hwp.SaveAs(os.path.join(path,name))
        time.sleep(0.2)
        self.hwp.Run("FileClose")
        time.sleep(0.2)


        self.hwp.XHwpWindows.Item(0).Visible = True
        self.hwp.Run("FileClose")
        time.sleep(0.2)

        self.hwp.Quit()


