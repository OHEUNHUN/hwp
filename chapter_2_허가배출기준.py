# -*- coding: utf-8 -*-
"""
Created on Tue Oct 12 16:13:49 2021

@author: ECOCNA_dev

chapter_2 허가배출기준 표 작성 프로그램 

"""

import win32com.client as win32
import os
import pandas as pd
import dbupdate_update
import time


class chapter_2_Premission:
   
   
    '''
        초기 설정으로 샘플 파일 경로, 샘플 파일 이름, 데이터베이스 이름, 출력될 파일 경로, 출력될 파일 이름 5가지 항목에 대한 값 설정필요.
        초기 값음 다음과 같이 설정이 되어 있음.
        
        self.samplepath = os.getcwd() - 파일이 설치된 위치
        self.filesample = "1_table_1.hwp"
        self.DbName = "hwp_1_10"
        self.filepath = "C:/Users/ECOCNA_dev/Desktop/testsave"
        self.fileName = "배출시설.hwp"
        
    '''


    def __init__(self):
        super().__init__()
        
        
        
        self.samplepath = os.getcwd() + '\sample'
        self.filesample = "2_table_2.hwp"
        
        self.DbName = "hwp_2_9"
        
        self.filepath = os.getcwd() + '\Result\chapter2_허가배출기준'
        self.fileName = "Chapter2_허가배출기준.hwp"

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
                self.filesample = "2_table_2.hwp"
                print("설정된 이름 없음")
        
        
        
            if DbName_in != "":
                self.DbName = DbName_in
                print("데이터베이스 이름 설정 성공")
            else:
                self.DbName = "hwp_2_9"
                print("설정된 이름 없음")
            
            
            
            if filePath_in != "":
                
                self.filepath = filePath_in
                print("파일이름 설정 성공")
                
            else:
                self.filepath = os.getcwd() + '\Result\chapter2_허가배출기준'
                print("설정된 이름 없음")
            
            
            if fileName_in != "":
                
                self.fileName = fileName_in
                print("파일이름 설정 성공")
                
            else:
                self.fileName = "Chapter2_허가배출기준.hwp"
                print("설정된 이름 없음")
                
            print("새로운 경로 입력 완료")  
        
        except :
            
            print("경로 설정 오류")
        


    def hwpfile_open(self):

        try :
            self.hwp.XHwpWindows.Item(0).Visible = True
            self.hwp.RegisterModule("FilePathCheckDLL","FilePathcheckerModule")
            
            file_name = self.filesample
            sample_path = self.samplepath
            
            
            self.hwp.Open(os.path.join(sample_path,file_name))
            self.hwp.Run("MoveDocBegin")
        
            self.hwp.Run("FileNew")
            self.hwp.XHwpWindows.Item(1).Visible = True
            
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

        num =0

        for i in pollusion_list:

            num +=1
            try :

                sel = {"황산화물": 3,"질소산화물":3,"일산화탄소" :2,"먼지":2,"아연화합물":2, "암모니아":2,"이황화탄소":2,
                "크롬화합물":2,"수은화합물":2, "구리화합물":2,"염화비닐":2, "황화수소":2 ,"디클로로메탄":2,"불소화합물":2,
                "페놀화합물":2,"포름알데히드":2}

                if sel[i] == 3:
                    self.hwp.MoveToField("배출구번호",select=False)
                    self.hwp.Run("TableInsertLowerRow")
                    self.hwp.MoveToField("단위시간",select=False)
                    self.hwp.Run("TableLowerCell")
                    self.hwp.Run("TableSplitCellRow2")
                    self.hwp.Run("TableSplitCellRow2")
                    self.hwp.MoveToField("단위량",select=False)
                    self.hwp.Run("TableLowerCell")
                    self.hwp.Run("TableSplitCellRow2")
                    self.hwp.Run("TableSplitCellRow2")
                    self.hwp.MoveToField("배출구번호",select=False)
                    self.hwp.Run("TableLowerCell")
                    self.hwp.Run("TableCellBlockRow")
                    self.hwp.Run("TableDistributeCellHeight")

                elif sel[i] ==2:
                    self.hwp.MoveToField("배출구번호",select=False)
                    self.hwp.Run("TableInsertLowerRow")
                    self.hwp.MoveToField("단위시간",select=False)
                    self.hwp.Run("TableLowerCell")
                    self.hwp.Run("TableSplitCellRow2")
                    self.hwp.MoveToField("단위량",select=False)
                    self.hwp.Run("TableLowerCell")
                    self.hwp.Run("TableSplitCellRow2")
                    self.hwp.MoveToField("배출구번호",select=False)
                    self.hwp.Run("TableLowerCell")
                    self.hwp.Run("TableCellBlockRow")
                    self.hwp.Run("TableDistributeCellHeight")

            except:
                    self.hwp.MoveToField("배출구번호",select=False)
                    self.hwp.Run("TableInsertLowerRow")

            #칸이 하나인 것과 셀병합을 하지 않기 위한 if문
            if num >1:
                self.hwp.MoveToField("배출구번호",select=True)
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
        self.inserttext(Select_data)



    def InsertText_chem(self, Select_data):
        table_list =["오염물질","농도단위","최대배출기준","허가배출기준","한계배출기준","예상배출농도","최저배출농도","참고기준","기준근거"]
        
        for i,data_list in enumerate(Select_data):

            for j, data in enumerate(data_list):

                self.hwp.MoveToField(table_list[j],select=False)

                for k in range(i+1):
                    self.hwp.Run("TableLowerCell")

                self.inserttext(data)

            print(data_list)
    
    
    def InsertText_time(self,Select_data):
        table_list = ["단위시간","단위량"]
        
        for i,data_list in enumerate(Select_data):

            for j, data in enumerate(data_list):

                self.hwp.MoveToField(table_list[j],select=False)

                for k in range(i+1):
                    self.hwp.Run("TableLowerCell")

                self.inserttext(data)

            print(data_list)





    def InsertCell(self,data_list):

        stack = data_list.drop_duplicates([0])
        stack = stack.sort_index(ascending=False)
        stack = stack[0].reset_index(drop=True)

        for i in stack:

            item = data_list.loc[data_list[0] == i]   
            item_chem = item.drop_duplicates([1])[1]
            item_chem = item_chem.reset_index(drop=True)

            self.MakeTableCol(item_chem)
            
            self.InsertText_stack(item[0].tolist()[0])


            item_chem_list = item.drop_duplicates([1])
            item_chem_list = item_chem_list.sort_index(ascending=False)
            item_chem_list = item_chem_list[[1,2,5,6,7,8,9,10,11]]
            item_chem_list = item_chem_list.values.tolist()

            self.InsertText_chem(item_chem_list)

            item_chem_time = item.sort_index(ascending=False)
            item_chem_time = item_chem_time[[3,4]]
            item_chem_time_list = item_chem_time.values.tolist()

            self.InsertText_time(item_chem_time_list)

            print(item)


        self.DeleteFristCell()





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

