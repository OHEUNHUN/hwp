# -*- coding: utf-8 -*-
"""
Created on Thu Oct 14 13:01:26 2021

@author: ECOCNA_dev
"""

import win32com.client as win32
import time
import os
import pandas as pd
import dbupdate_update





class chapter_4_protect:
    
    
    def __init__(self):
        
        '''
        초기 설정으로 샘플 파일 경로, 샘플 파일 이름, 데이터베이스 이름, 출력될 파일 경로, 출력될 파일 이름 6가지 항목에 대한 값 설정필요.
        초기 값음 다음과 같이 설정이 되어 있음.
        
        self.samplepath = os.getcwd() - 파일이 설치된 위치
        self.filesample = "4_table_2.hwp"
        self.DbName_f = "hwp_4_4_f"
        self.DbName_p = "hwp_4_4_p"
        self.filepath = "C:/Users/ECOCNA_dev/Desktop/testsave"
        self.fileName = "배출시설.hwp"
        '''
        super().__init__()
        
        self.samplepath = os.getcwd() + '\sample'
        self.filesample = "4_table_2.hwp"
        
        self.DbName_f = "hwp_4_4_f"
        self.DbName_p = "hwp_4_4_p"
        
        self.filepath = os.getcwd() + '\Result\chapter4_방지시설'
        self.fileName = "Chapter4_방지시설.hwp"
        
        
        
    def Get_property(self):
        '''
        초기 설정된 6개 항목에 대한 값을 불러 올 수 있다.
        
        '''
        return self.samplepath, self.filesample, self.DbName_f, self.DbName_p ,self.filepath, self.fileName
    
    
    
    def Set_property(self,samplepath_in,filesample_in,DbName_f_in,DbName_p_in ,filePath_in ,fileName_in):
        '''
        초기 설정을 변경할 수 있다.
        변경할 필요가 없다면 Enter를 쳐서 넘길 수 있다. 
        변경에 성공했다면 설정 성공리나는 말이 출력된다.
        실패하면 **없음으로 출력된다
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
                self.filesample = "4_table_2.hwp"
                print("설정된 이름 없음")
        
        
        
            if DbName_f_in != "":
                self.DbName_f = DbName_f_in
                print("데이터베이스 이름 설정 성공")
            else:
                self.DbName_f = "hwp_4_4_f"
                print("설정된 이름 없음")
            
            
            if DbName_p_in != "":
                self.DbName_p = DbName_p_in
                print("데이터베이스 이름 설정 성공")
            else:
                self.DbName_p = "hwp_4_4_p"
                print("설정된 이름 없음")
            
            
            
            if filePath_in != "":
                
                self.filepath = filePath_in
                print("파일이름 설정 성공")
                
            else:
                self.filepath = os.getcwd() + '\Result\chapter4_방지시설'
                print("설정된 이름 없음")
            
            
            if fileName_in != "":
                
                self.fileName = fileName_in
                print("파일이름 설정 성공")
                
            else:
                self.fileName = "Chapter4_방지시설.hwp"
                print("설정된 이름 없음")
                
            print("새로운 경로 입력 완료")  
        
        except :
            
            print("경로 설정 오류")
        
        
    
        
    def hwpfile_open(self):
        
        '''
        파일 이름에는 .hwp가 붙을 것 이후 수정하기
        
        파일 패스 생성 및 파일 이름으로 한글 파일 열기
        
        파일이 열리면 새로운 한글 파일 2개가 추가로 열림
        '''
        
        try:
            self.hwp.XHwpWindows.Item(0).Visible = False
            'Visible True라고 하면 창이 보이고 False면 백그라운드 실행'
            self.hwp.RegisterModule("FilePathCheckDLL","FilePathcheckerModule")
        
        
            file_name = self.filesample
            sample_path = self.samplepath
        
        
            self.hwp.Open(os.path.join(sample_path,file_name))
        
            '패스 경로 및 파일 이름을 선택해서 읽는다'
        
    
            self.hwp.Run("MoveDocBegin")
            self.hwp.Run("FileNew")
            self.hwp.XHwpWindows.Item(1).Visible = False
            self.hwp.Run("FileNew")
            self.hwp.XHwpWindows.Item(2).Visible = False
        
        
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

    
    def Copy_Update_table(self):
        '''
        만들어진 표를 새문서에 복사해서 붙여 넣는다.
        '''
        

        self.hwp.XHwpDocuments.Item(1).SetActive_XHwpDocument()
        self.hwp.Run("CloseEx")
        self.hwp.Run("SelectAll")
        self.hwp.Run("Copy")
        self.hwp.XHwpDocuments.Item(2).SetActive_XHwpDocument()
        self.hwp.Run("MoveDocEnd")
        self.hwp.Run("Paste")


    def GetProcessType(self,DataFrameName_f,Num):
        '''공정 대분류 Num: 1,공정 중분류 Num: 2 
            대분류는 00-##-@@에서 00을 가지고 분류를 한다.
            중분류는 00-##-@@에서 00-##이 같은 단위로 분류를 한다.
        ''' 
        
        if Num == 1:
            
            gettype = list(set(DataFrameName_f[13].str[:2]))
            gettype.sort()
            
            return gettype
            
        elif Num==2:
            
            gettype = list(set(DataFrameName_f[13].str[:5]))
            gettype.sort()
            
            return gettype
            
        else:
            
            gettype = list(set(DataFrameName_f[13].str[:5]))
            gettype.sort()
            
            return gettype


    def GetProcessList(self,DataFrameName_f,processType,Num):
        
        '''공정 대분류 Num: 1,공정 중분류 Num: 2 
           GetProcessType과 같은 숫자를 사용해야한다.
        ''' 
        
        
        process_list=[[0]]*int(len(processType))
        
        
        if Num == 1:
            for i in range(0,int(len(processType))):

                process_list[i] = DataFrameName_f[DataFrameName_f[13].str[:2] == processType[i]]
                '대공정 종류별로 select문을 만들어서 물질을 부분을 받아온다 (가장 앞에 2개 숫자 **-00-00)'
            
            return process_list
        
        
        elif Num == 2:
            
            for i in range(0,int(len(processType))):
                
                process_list[i] = DataFrameName_f[DataFrameName_f[13].str[:5] == processType[i]]
                '중간 공정 종류별로 select문을 만들어서 물질을 부분을 받아온다()'
                
                
            return process_list



    def inserttext(self,text):
        
        '''
        전체선택이 포함된 상태에서 한글 파일에 글자가 입력됨. 
        기존에 남아 있는 것을 지우고 덮어쓰는 역할
        '''
        
        self.hwp.Run("SelectAll")
        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
        self.hwp.HParameterSet.HInsertText.Text = text
        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)



    def inserttext_notall(self,text):
        
        '''
        일반적인 글자 입력
        커서가 위치한 곳부터 글자가 입력
        '''
        
        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
        self.hwp.HParameterSet.HInsertText.Text = text
        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)





    def InsertTbl_f(self, DataFrameName, DataNum):
        
        '''
        한글표에 있는 필드이름에 데이터를 넣는다.
        '''

        print(DataFrameName)
        self.hwp.PutFieldText("방지시설관리번호",DataFrameName.loc[DataNum][0])
        self.hwp.PutFieldText("시설명",DataFrameName.loc[DataNum][1])
        self.hwp.PutFieldText("용량",DataFrameName.loc[DataNum][2])
        self.hwp.PutFieldText("용량단위",DataFrameName.loc[DataNum][3])
        self.hwp.PutFieldText("수량",DataFrameName.loc[DataNum][4])
        self.hwp.PutFieldText("일일가동시간",DataFrameName.loc[DataNum][5])
        self.hwp.PutFieldText("연간가동일수",DataFrameName.loc[DataNum][6])
        self.hwp.PutFieldText("압력손실",DataFrameName.loc[DataNum][7])
        text_data = DataFrameName.loc[DataNum][8].split(',')
        self.hwp.MoveToField("전단시설")
        for i in range(0,int(len(text_data))):
            self.inserttext_notall(str(text_data[i])+str(',')+str('\r'))
            if i==int(len(text_data)-1):
                self.hwp.Run("DeleteBack")
                self.hwp.Run("DeleteBack")


        text_data = DataFrameName.loc[DataNum][9].split(',')
        self.hwp.MoveToField("후단시설")
        for i in range(0,int(len(text_data))):
            self.inserttext_notall(str(text_data[i])+str(',')+str('\r'))
            if i==int(len(text_data)-1):
                self.hwp.Run("DeleteBack")
                self.hwp.Run("DeleteBack")

        self.hwp.PutFieldText("부대시설정보",DataFrameName.loc[DataNum][10])
        self.hwp.PutFieldText("비고",DataFrameName.loc[DataNum][11])





    def searchTbl_p(self,DataFrameName_p,DataFrameName_f,Num):
        
        '''
        먼저 입력된 xx_F 뷰에 있는 정보를 기준으로 xx_P에 대한 정보를 가지고 온다.
        return 값은 contrl_type_list,contrl_type,Select_license_num
    
        contrl_type_list : 리스트 형태의 데이터를 반환하며, select문으로 선택 된 데이터에서 license부분의 가장 앞에 있는 오염물질의 종류를 
        중복에 상관하지 않고 모두 반환한다.['A','A','A','W','W','W' ]
        
        contrl_type : 위에 받은 리스트에서 중복값은 뺀 순수 종류에 대한 리스트를 반환한다. ['A','W']
        
        
        Select_license_num : 선택된 데이터 전체를 반환한다.
    
        '''
        
        
        Select_license_num = DataFrameName_p.loc[DataFrameName_p[14] == DataFrameName_f[12][Num]]
        
        Select_license_num = Select_license_num.reset_index()
        'select된 데이터에서 index번호를 초기화 한다.아니면 이전 번호가 계속 따라오게됨.'
        
        contrl_type_list=[]
        
        for i in range(0,int(len(Select_license_num))):
            
            contrl_type_list.append(Select_license_num[0][i].split("-")[0])
            
        contrl_type= list(set(contrl_type_list))
        
        return contrl_type_list, contrl_type, Select_license_num





    def DelTbl_p(self,input_type, input_type_list):
        '''
        input_type : 오염물질의 종류를 받아온다. ex:["AT","OT","VT","NT","WT","ST","NpR","FpT","FfT","PT"]
        
        input_type_list : 오염물질이 종류별로 얼마나 있는지 파악을 한다.

        '''
        Data_type = input_type
        Data_list = input_type_list
        
        maplist_up = ["S2","S3","S4","S5","S6","S7","S8","S9","S10","S11"]
        maplist_down = ["A14","A15","A16","A17","A18","A19","A20","A21","A22","A23"]
        MapListType =["AT","OT","VT","NT","WT","ST","NpR","FpT","FfT","PT"]
        
        Dic_type_up = dict(zip(MapListType,maplist_up))
        Dic_type_down = dict(zip(MapListType,maplist_down))
    
    
        Delet_Tbl_type = MapListType
        
        
        for i in range(0,int(len(Data_type))):
            Delet_Tbl_type.remove(Data_type[i])
        
        
        if len(Data_list) == 0:
            for i in range(1,int(len(Delet_Tbl_type))):
                self.hwp.MoveToField(str(Dic_type_up[Delet_Tbl_type[i]]))
                self.hwp.Run("TableDeleteRow")
                self.hwp.MoveToField(str(Dic_type_down[Delet_Tbl_type[i]]))
                self.hwp.Run("TableDeleteRow")
        else :
            for i in range(0,int(len(Delet_Tbl_type))):
                self.hwp.MoveToField(str(Dic_type_up[Delet_Tbl_type[i]]))
                self.hwp.Run("TableDeleteRow")
                self.hwp.MoveToField(str(Dic_type_down[Delet_Tbl_type[i]]))
                self.hwp.Run("TableDeleteRow")




    def MakeTableCol_p(self,input_type,input_type_list):
        
        '''
        오염 물질의 종류와 종류별 숫자 리스트를 불러온다
        
        오염물질 종류에 따른 위치 좌표를 딕셔너리 형태로 호출한다. 
        
        count를 이용하여 안에 있는 물질 종류별 갯수를 파악한다. 
        
        이후 해당 셀로 이동하여 행을 추가하고 병합하는 것을 반복한다.
        '''
        
        Data_type = input_type
        Data_list = input_type_list
        
        
        maplist_up = ["S2","S3","S4","S5","S6","S7","S8","S9","S10","S11"]
        maplist_down = ["A14","A15","A16","A17","A18","A19","A20","A21","A22","A23"]
        maplist_flux_unit = ["unit14","unit15","unit16","unit17","unit18","unit19","unit20","unit21","unit22","unit23"]
        maplist_temp = ["temp14","temp15","temp16","temp17","temp18","temp19","temp20","temp21","temp22","temp23"]
        MapListType =["AT","OT","VT","NT","WT","ST","NpR","FpT","FfT","PT"]
        
        Dic_type_up = dict(zip(MapListType,maplist_up))
        Dic_type_down = dict(zip(MapListType,maplist_down))
        Dic_unit =  dict(zip(MapListType,maplist_flux_unit))
        Dic_temp = dict(zip(MapListType,maplist_temp))
    
        
        '카운트는 딕셔러닐 형태로 해당되는 키가 있으면 data_list별로 몇개가 있는지 확인이 가능하다.'
        
        count={}
        for i in Data_list:
            try: count[i] += 1
            except: count[i]=1
        
        
        '행을 추가하고 병합하는 것을 반복한다. 항목별 병합이 끝나면 시설명과 배출관리 변호를 마지막에 통합한다.'
        for i in range(0,int(len(Data_type))):
            
            for j in range(0,int(count[Data_type[i]])-1):
                self.hwp.MoveToField(str(Dic_type_up[Data_type[i]]))
                self.hwp.Run("TableInsertLowerRow")
                self.hwp.MoveToField(str(Dic_type_down[Data_type[i]]))
                self.hwp.Run("TableInsertLowerRow")
                self.hwp.MoveToField(str(Dic_type_up[Data_type[i]]), select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField(str(Dic_type_down[Data_type[i]]), select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField(str(Dic_unit[Data_type[i]]), select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField(str(Dic_temp[Data_type[i]]), select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")




                self.hwp.MoveToField("방지시설관리번호", select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField("시설명", select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField("용량", select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField("용량단위", select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField("수량", select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField("일일가동시간", select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField("연간가동일수", select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                self.hwp.MoveToField("압력손실", select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                
        self.hwp.MoveToField("전단시설", select = True)
        self.hwp.Run("TableCellBlockExtend")
        self.hwp.Run("TableColPageDown")
        self.hwp.Run("TableMergeCell")
        self.hwp.MoveToField("후단시설", select = True)
        self.hwp.Run("TableCellBlockExtend")
        self.hwp.Run("TableColPageDown")
        self.hwp.Run("TableMergeCell")
        self.hwp.MoveToField("부대시설정보", select = True)
        self.hwp.Run("TableCellBlockExtend")
        self.hwp.Run("TableColPageDown")
        self.hwp.Run("TableMergeCell")
        self.hwp.MoveToField("비고", select = True)
        self.hwp.Run("TableCellBlockExtend")
        self.hwp.Run("TableColPageDown")
        self.hwp.Run("TableMergeCell")
        
        return count

    

    def InsertTbl_p(self,input_Data_list, input_type_list, count):
        
        
        Data_list_ori= input_Data_list
        type_list = input_type_list
        
        maplist_up = ["S2","S3","S4","S5","S6","S7","S8","S9","S10","S11"]
        maplist_down = ["A14","A15","A16","A17","A18","A19","A20","A21","A22","A23"]
        MapListType =["AT","OT","VT","NT","WT","ST","NpR","FpT","FfT","PT"]
        
        Dic_type_up = dict(zip(MapListType,maplist_up))
        Dic_type_down = dict(zip(MapListType,maplist_down))
        
        Data_list_ori.fillna('-',inplace=True)
        
        for K in range(0,int(len(type_list))):
            

            Data_list = Data_list_ori[Data_list_ori[0].str.contains(type_list[K])]
            Data_list = Data_list.reset_index()


            for i in range(0,int(count[type_list[K]])):

                self.hwp.PutFieldText(Dic_type_up[type_list[K]],Data_list.loc[i][0])
                self.hwp.MoveToField(Dic_type_up[type_list[K]])
                self.hwp.Run("TableRightCellAppend")
                
                for j in range(0,i):
                    self.hwp.Run("TableLowerCell")
    
                self.inserttext(Data_list.loc[i][1])
                self.hwp.Run("TableRightCellAppend")
                
                if int(Data_list.loc[i][2]) < 1 : 
                    self.inserttext(Data_list.loc[i][2]*100)
                    
                else :
                    self.inserttext(Data_list.loc[i][2])


                '아래쪽으로 이동함'
                self.hwp.MoveToField(Dic_type_down[type_list[K]])


                try :
                    self.inserttext((format(Data_list.loc[i][3],'10.2E')))
                except :
                    self.inserttext((format(Data_list.loc[i][3])))

                self.hwp.Run("TableRightCellAppend")

                self.inserttext(str(Data_list.loc[i][4]))
                self.hwp.Run("TableRightCellAppend")


                for j in range(0,i):
                    self.hwp.Run("TableLowerCell")


                try :
                    self.inserttext((format(Data_list.loc[i][5],'10.2E')))
                except :
                    self.inserttext((format(Data_list.loc[i][5])))

                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][6]))
                self.hwp.Run("TableRightCellAppend")

                try :
                    self.inserttext((format(Data_list.loc[i][7],'10.2E')))
                except :
                    self.inserttext((format(Data_list.loc[i][7])))

                self.hwp.Run("TableRightCellAppend")

                try :
                    self.inserttext((format(Data_list.loc[i][8],'10.2E')))
                except :
                    self.inserttext((format(Data_list.loc[i][8])))

                self.hwp.Run("TableRightCellAppend")

                try :
                    self.inserttext((format(Data_list.loc[i][9],'10.2E')))
                except :
                    self.inserttext((format(Data_list.loc[i][9])))

                self.hwp.Run("TableRightCellAppend")

                self.inserttext(str(Data_list.loc[i][10]))
                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][11]))
                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][12]))
                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][13]))





    def Main(self,host,port,database,user,password):
                
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        dbname_f = self.DbName_f
        'hwp_4_4_f'
        dbname_p = self.DbName_p
        'hwp_4_4_p'
        
   
        db = dbupdate_update.DBCONN()
        db.SetDataBase(host,port,database,user,password)
        db.DB_CONN()
   
    
        hwp_4_4_f = db.query (dbname_f,"*")
        hwp_4_4_f = pd.DataFrame(hwp_4_4_f)

        hwp_4_4_p =db.query (dbname_p,"*")
        hwp_4_4_p = pd.DataFrame(hwp_4_4_p)


        gettype = self.GetProcessType(hwp_4_4_f,2)


        hwp_4_4_f[8] = hwp_4_4_f[8].str.replace(',',',\n')
        hwp_4_4_f[9] = hwp_4_4_f[9].str.replace(',',',\n')

        Manufamaplist = ["PU","P","PW"]
        ManufaType =["01","02","03"]        
        Manufa_type = dict(zip(ManufaType,Manufamaplist))


        self.hwpfile_open()
        path = os.path.abspath(self.filepath)



        for i in range(0,int(len(gettype))):
    
            self.Copy_Table_sample()
            hwp_table = db.query_like(dbname_f,"*",gettype[i])
            hwp_table = pd.DataFrame(hwp_table)
            hwp_table.fillna('-',inplace=True)


            self.hwp.Run("MoveDocBegin")
            text = hwp_table[13][0].split("-")
            name = Manufa_type[text[0]]+"-"+text[1]+"-"+text[2]
            self.inserttext_notall(name)

    
            for j in range(0,int(len(hwp_table))):
        
                contrl_type_list, contrl_type, Select_data = self.searchTbl_p(hwp_4_4_p, hwp_table, j)

                Data_type = contrl_type
                Data_list = contrl_type_list

    
                try:
                    self.DelTbl_p(Data_type, Data_list)
                    type_count =  self.MakeTableCol_p(Data_type, Data_list)
                    self.InsertTbl_f(hwp_table, j)
                    self.InsertTbl_p(Select_data, Data_type, type_count)
        
                except:
                    pass

                self.Copy_Update_table()
                self.Copy_Table_sample()

        

            name = name + '_4장_' + self.fileName
            
            self.hwp.XHwpDocuments.Item(2).SetActive_XHwpDocument()
            time.sleep(0.1)
            self.hwp.SaveAs(os.path.join(path,name))
            time.sleep(0.1)
            self.hwp.Run("FileClose")
            time.sleep(0.1)
            self.hwp.Run("FileNew")
            self.hwp.XHwpWindows.Item(2).Visible = False
            time.sleep(0.1)

        self.hwp.XHwpWindows.Item(2).Visible = True
        self.hwp.Run("FileClose")
        time.sleep(0.1)
        self.hwp.XHwpWindows.Item(1).Visible = True
        self.hwp.SaveAs(os.path.join(path,'trash_4장방지시설.hwp'))
        self.hwp.Run("FileClose")
        time.sleep(0.1)
        self.hwp.XHwpWindows.Item(0).Visible = True
        self.hwp.Run("FileClose")
        time.sleep(0.1)


        self.hwp.Quit()





