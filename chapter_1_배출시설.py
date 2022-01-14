# -*- coding: utf-8 -*-
"""
Created on Tue Oct 26 12:06:51 2021

@author: ECOCNA_dev
"""

import os
import pandas as pd
import win32com.client as win32
import dbupdate_update
import time


class chapter_1_production:
    
    '''
    배출시설 출력에 대한 파일 
    
    method 목록 :
        __init__  초기값 설정
        Get_property
        Set_property
        hwpfile_open
        Copy_Table_sample
        inserttext_notall
        InsertCell
    '''
    
    def __init__(self):
        '''
        초기 설정으로 샘플 파일 경로, 샘플 파일 이름, 데이터베이스 이름, 출력될 파일 경로, 출력될 파일 이름 5가지 항목에 대한 값 설정필요.
        초기 값음 다음과 같이 설정이 되어 있음.
        
        self.samplepath = os.getcwd() - 파일이 설치된 위치
        self.filesample = "1_table_1.hwp"
        self.DbName = "hwp_1_10"
        self.filepath = "C:/Users/ECOCNA_dev/Desktop/testsave"
        self.fileName = "배출시설.hwp"
        
        '''
        super().__init__()
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        self.samplepath = os.getcwd() + '\sample'
        self.filesample = "1_table_1.hwp"
        
        self.DbName = "hwp_1_10"
        
        self.filepath = os.getcwd() + '\Result\chapter1_배출시설'
        self.fileName = "Chapter1_배출시설.hwp"
        
        
        
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
                self.filesample = "1_table_1.hwp"
                print("설정된 이름 없음")
        
        
        
            if DbName_in != "":
                self.DbName = DbName_in
                print("데이터베이스 이름 설정 성공")
            else:
                self.DbName = "hwp_1_10"
                print("설정된 이름 없음")
            
            
            
            if filePath_in != "":
                
                self.filepath = filePath_in
                print("파일이름 설정 성공")
                
            else:
                self.filepath = os.getcwd() + '\Result\chapter1_배출시설'
                print("설정된 이름 없음")
            
            
            if fileName_in != "":
                
                self.fileName = fileName_in
                print("파일이름 설정 성공")
                
            else:
                self.fileName = "Chapter1_배출시설.hwp"
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



    def inserttext_notall(self,text):
        
        '''
        일반적인 글자 입력
        커서가 위치한 곳부터 글자가 입력
        
        '''
        
        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
        self.hwp.HParameterSet.HInsertText.Text = text
        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)



    def InsertCell(self,data_list):
        
        '''
        데이터에 있는 파일을 양식에 맞게 입력한다. 
        
        양식이 변하면 반드시 수정해야함 
        
        한글 표에 있는 필드이름의 위치는 절대로 바꾸면 안됨. 
        
        관리 번호로 이동한다. 
        
        '''
        self.hwp.MoveToField("관리번호")
        
        data= data_list

        print(data)
        
        for i in range(0,int(len(data))):
            for j in range(0,int(len(data.loc[i]))):
                if j == 8:
                    text_data = data.loc[i][j].split(",")
                    '배출 물질을 .을 기준으로 자른다'
                    for k in range(0,int(len(text_data))):
                        self.inserttext_notall(str(text_data[k])+str(',')+str('\r'))
                        '입력할때 잘린 배출물질의 갯수만큼 추가 한다'
                        if k==int(len(text_data)-1):
                            
                            self.hwp.Run("DeleteBack")
                            self.hwp.Run("DeleteBack")
                            '마지막에 아래로 내려가는 줄은 지운다'
                else:
                    self.inserttext_notall(str(data.loc[i][j]))
                self.hwp.Run("TableRightCellAppend")
                '8번째가 아니면 옆으로 한칸씩 이용하며 값을 넣는다'
            if i == int(len(data)-1) :
                self.hwp.Run("TableDeleteRow")
                '마지막 열에 도달하면 아래로 한줄을 내린다'
             
        self.hwp.Run("CloseEx")





    def Main(self,host,port,database,user,password):
        
        dbname = self.DbName
        
        db = dbupdate_update.DBCONN()
        db.SetDataBase(host,port,database,user,password)
        db.DB_CONN()
        
        hwp_table = db.query(dbname, "*")
        hwp_table = pd.DataFrame(hwp_table)
        hwp_table.fillna('-',inplace=True)
        
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




