# -*- coding: utf-8 -*-
"""
Created on Tue Oct 19 13:02:49 2021

@author: ECOCNA_dev
"""

import win32com.client as win32
import time
import os
import pandas as pd
import dbupdate_update




class chapter_4_production:
    
    
    def __init__(self):
        super().__init__()
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        self.samplepath = os.getcwd() + '\sample'
        self.filesample = "4_table_1.hwp"
        
        self.DbName_f = "hwp_4_3_f"
        self.DbName_p = "hwp_4_3_p"
        
        self.filepath = os.getcwd() + '\Result\chapter4_배출시설'
        self.fileName = "Chapter4_배출시설.hwp"


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
                self.samplepath = os.getcwd() + '\sample'
                print("설정된 경로 없음")
        
        
        
            if filesample_in != "":
                self.filesample = filesample_in
                print("파일샘플이름 설정 성공")
            else:
                self.filesample = "4_table_1.hwp"
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
                self.filepath = os.getcwd() + '\Result\chapter4_배출시설'
                print("설정된 이름 없음")
            
            
            if fileName_in != "":
                
                self.fileName = fileName_in
                print("파일이름 설정 성공")
                
            else:
                self.fileName = "Chapter4_배출시설.hwp"
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
            
    
        
        
        
    def Copy_Table_sample(self):
        
        
        
        
        self.hwp.XHwpDocuments.Item(0).SetActive_XHwpDocument()
        self.hwp.Run("SelectAll")
        self.hwp.Run("Copy")
        self.hwp.XHwpDocuments.Item(1).SetActive_XHwpDocument()
        self.hwp.Run("SelectAll")
        self.hwp.Run("Paste")
        
        '''self.hwp.Run("FileClose")'''
      
    
    def Copy_Update_table(self):
        
        
        
        
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
            
            gettype = list(set(DataFrameName_f[18].str[:2]))
            gettype.sort()
            
            return gettype
            
        elif Num==2:
            
            gettype = list(set(DataFrameName_f[18].str[:5]))
            gettype.sort()
            
            return gettype
            
        else:
            
            gettype = list(set(DataFrameName_f[18].str[:5]))
            gettype.sort()
            
            return gettype
            
        
        
    def GetProcessList(self,DataFrameName_f,processType,Num):
        
        '''
        공정 대분류 Num: 1,공정 중분류 Num: 2 
        GetProcessType과 같은 숫자를 사용해야한다.
        ''' 
        
        
        process_list=[[0]]*int(len(processType))
        
        
        if Num == 1:
            for i in range(0,int(len(processType))):

                process_list[i] = DataFrameName_f[DataFrameName_f[18].str[:2] == processType[i]]
            
            return process_list
        
        
        elif Num == 2:
            
            for i in range(0,int(len(processType))):
                
                process_list[i] = DataFrameName_f[DataFrameName_f[18].str[:5] == processType[i]]
            
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

    



    def InsertTbl_f(self,DataFrameName,DataNum):
        '''
        한글표에 있는 필드이름에 데이터를 넣는다.
        '''

        print(DataFrameName)

        self.hwp.PutFieldText("A2",DataFrameName.loc[DataNum][0])
        self.hwp.PutFieldText("B2",DataFrameName.loc[DataNum][1])
        self.hwp.PutFieldText("C2",DataFrameName.loc[DataNum][2])
        self.hwp.PutFieldText("E2",DataFrameName.loc[DataNum][3])
        self.hwp.PutFieldText("F2",DataFrameName.loc[DataNum][4])
        self.hwp.PutFieldText("G2",DataFrameName.loc[DataNum][5])
        self.hwp.PutFieldText("I2",DataFrameName.loc[DataNum][6])
        self.hwp.PutFieldText("J2",DataFrameName.loc[DataNum][7])
        self.hwp.PutFieldText("L2",DataFrameName.loc[DataNum][8])

        text_data = DataFrameName.loc[DataNum][9].split(',')
        self.hwp.MoveToField("N2")
        for i in range(0,int(len(text_data))):
            self.inserttext_notall(str(text_data[i])+str(',')+str('\r'))
            if i==int(len(text_data)-1):
                self.hwp.Run("DeleteBack")
                self.hwp.Run("DeleteBack")


        text_data = DataFrameName.loc[DataNum][10].split(',')
        self.hwp.MoveToField("M2")
        for i in range(0,int(len(text_data))):
            self.inserttext_notall(str(text_data[i])+str(',')+str('\r'))
            if i==int(len(text_data)-1):
                self.hwp.Run("DeleteBack")
                self.hwp.Run("DeleteBack")

        self.hwp.PutFieldText("P2",DataFrameName.loc[DataNum][11])
        self.hwp.PutFieldText("R2",DataFrameName.loc[DataNum][12])
        self.hwp.PutFieldText("T2",DataFrameName.loc[DataNum][13])
        self.hwp.PutFieldText("U2",DataFrameName.loc[DataNum][14])
        self.hwp.PutFieldText("V2",DataFrameName.loc[DataNum][15])
        self.hwp.PutFieldText("W2",DataFrameName.loc[DataNum][16])      


    
    
    def searchTbl_p(self,DataFrameName_p,DataFrameName_f,Num):
        
        '''
        먼저 입력된 xx_F 뷰에 있는 정보를 기준으로 xx_P에 대한 정보를 가지고 온다.
        return 값은 contrl_type_list,contrl_type,Select_license_num
    
        contrl_type_list : 리스트 형태의 데이터를 반환하며, select문으로 선택 된 데이터에서 license부분의 가장 앞에 있는 오염물질의 종류를 
        중복에 상관하지 않고 모두 반환한다.['A','A','A','W','W','W' ]
        
        contrl_type : 위에 받은 리스트에서 중복값은 뺀 순수 종류에 대한 리스트를 반환한다. ['A','W']
        
        
        Select_license_num : 선택된 데이터 전체를 반환한다.
    
        '''
        
        Select_license_num = DataFrameName_p.loc[DataFrameName_p[14] == DataFrameName_f[17][Num]]
        
        Select_license_num = Select_license_num.reset_index()
        'select된 데이터에서 index번호를 초기화 한다.아니면 이전 번호가 계속 따라오게됨.'
        
        contrl_type_list=[]
        
        
        
        for i in range(0,int(len(Select_license_num))):
            
            contrl_type_list.append(Select_license_num[0][i].split("-")[0])
            
        contrl_type= list(set(contrl_type_list))
        
        
        return contrl_type_list, contrl_type, Select_license_num   
    



    def DelTbl_p(self,input_type, input_type_list):
        '''
        input_type : 오염물질의 종류를 받아온다. ex:["A","O","V","N","W","S","Np","Ws","WsD","Fp","Ff","P"]
        
        input_type_list : 오염물질이 종류별로 얼마나 있는지 파악을 한다.

        '''
        
        Data_type = input_type
        
        maplist = ["C6","C8","C10","C12","C14","C16","C18","C20","C22","C24","C26","C28"]
        MapListType =["A","O","V","N","W","S","Np","Ws","WsD","Fp","Ff","P"]
        
        Dic_type = dict(zip(MapListType,maplist))
    
        Delet_Tbl_type = MapListType
        
        for i in range(0,int(len(Data_type))):
            Delet_Tbl_type.remove(Data_type[i])
        
        
        for i in range(0,int(len(Delet_Tbl_type))):
            self.hwp.MoveToField(str(Dic_type[Delet_Tbl_type[i]]))
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
        
        maplist = ["C6","C8","C10","C12","C14","C16","C18","C20","C22","C24","C26","C28"]
        MapListType =["A","O","V","N","W","S","Np","Ws","WsD","Fp","Ff","P"]
        
        Dic_type = dict(zip(MapListType,maplist))
        
        
        '카운트는 딕셔러닐 형태로 해당되는 키가 있으면 data_list별로 몇개가 있는지 확인이 가능하다.'
        
        count={}
        for i in Data_list:
            try: count[i] += 1
            except: count[i]=1
        
        
        '행을 추가하고 병합하는 것을 반복한다. 항목별 병합이 끝나면 시설명과 배출관리 변호를 마지막에 통합한다.'
        for i in range(0,int(len(Data_type))):
            for j in range(0,int(count[Data_type[i]])-1):
                self.hwp.MoveToField(str(Dic_type[Data_type[i]]))
                self.hwp.Run("TableInsertLowerRow")
                self.hwp.MoveToField(str(Dic_type[Data_type[i]]), select = True)
                self.hwp.Run("TableCellBlockExtend")
                self.hwp.Run("TableLowerCell")
                self.hwp.Run("TableMergeCell")
                
        self.hwp.MoveToField("A2", select = True)
        self.hwp.Run("TableCellBlockExtend")
        self.hwp.Run("TableColPageDown")
        self.hwp.Run("TableMergeCell")
        self.hwp.MoveToField("B2", select = True)
        self.hwp.Run("TableCellBlockExtend")
        self.hwp.Run("TableColPageDown")
        self.hwp.Run("TableMergeCell")


        return count


    def InsertTbl_p_word(self, text_a, text_b):

        MapListType =["A","O","V","N","W","S","Np","Ws","WsD","Fp","Ff","P"]
        try:
            text_a = round(float(text_a),3)
            text_b = round(float(text_b),3)
        except:
            pass

        textformat =["■ 대기 - (배출유량 : {} Sm³/min, 온도 : {}°C, 발생농도 단위 : (입자상) mg/Sm³ (가스상) ppm)".format(text_a,text_b),
                    "■ 악취",
                    "■ 휘발성유기화합물",
                    "■ 소음·진동",
                    "■ 수질 (유량 : {} 톤/일,    배출온도 : {}°C)".format(text_a,text_b),
                    "■ 토양",
                    "■ 비점오염원",
                    "■ 폐기물",
                    "■ 폐기물처리",
                    "■ 비산먼지 - (배출유량 : {} Sm³/min, 온도 : {}°C, 발생농도 단위 : (입자상) mg/Sm³ (가스상) ppm)".format(text_a,text_b),
                    "■ 비산배출",
                    "■ 잔류성유기오염물질"]


        Dic_format_list = dict(zip(MapListType,textformat))


        return Dic_format_list



    def InsertTbl_p(self,input_Data_list, input_type_list, count):
        
        Data_list_ori= input_Data_list
        type_list = input_type_list
        Data_list_ori.fillna('-',inplace=True)
        
        maplist = [["C6","D6","D7"],["C8","D8","D9"],["C10","D10","D11"],
           ["C12","D12","D13"],["C14","D14","D15"],["C16","D16","D17"],
           ["C18","D18","D19"],["C20","D20","D21"],["C22","D22","D23"],
           ["C24","D24","D25"],["C26","D26","D27"],["C28","D28","D29"]]
        MapListType =["A","O","V","N","W","S","Np","Ws","WsD","Fp","Ff","P"]
        
        Dic_type = dict(zip(MapListType,maplist))
        
        for k in range(0,int(len(type_list))):
           
        
            Data_list = Data_list_ori[Data_list_ori[0].str.contains(type_list[k])]
            Data_list = Data_list.reset_index()


            for i in range(0,int(count[type_list[k]])):


                self.hwp.PutFieldText(Dic_type[type_list[k]][0],Data_list.loc[i][0])


                if len(Data_list.loc[i][15]) == 2:
                    
                    if Data_list.loc[i][0].split('-')[0] == 'A':
                        self.hwp.PutFieldText(Dic_type[type_list[k]][1],'■ 대기' )

                    elif Data_list.loc[i][0].split('-')[0] == 'Fp':
                        self.hwp.PutFieldText(Dic_type[type_list[k]][1],'■ 비산먼지')
                    else :
                        text_word= self.InsertTbl_p_word(str(Data_list.loc[i][1]),str(Data_list.loc[i][2]))
                        self.hwp.PutFieldText(Dic_type[type_list[k]][1],text_word[type_list[k]])
                elif Data_list.loc[i][15].split(',')[1] == '01':
                    
                    if Data_list.loc[i][0].split('-')[0] == 'A':
                        self.hwp.PutFieldText(Dic_type[type_list[k]][1],'■ 대기 - ' + Data_list.loc[i][16])

                    elif Data_list.loc[i][0].split('-')[0] == 'Fp':
                        self.hwp.PutFieldText(Dic_type[type_list[k]][1],'■ 비산먼지 - '+ Data_list.loc[i][16])

                    else :
                        text_word= self.InsertTbl_p_word(str(Data_list.loc[i][1]),str(Data_list.loc[i][2]))
                        self.hwp.PutFieldText(Dic_type[type_list[k]][1],text_word[type_list[k]])
                else :

                    text_word= self.InsertTbl_p_word(str(Data_list.loc[i][1]),str(Data_list.loc[i][2]))
                    self.hwp.PutFieldText(Dic_type[type_list[k]][1],text_word[type_list[k]])
                


                self.hwp.MoveToField(Dic_type[type_list[k]][2],select=True)


                for j in range(0,i):
                    self.hwp.Run("TableLowerCell")
            

                Data_list.loc[i] = Data_list.loc[i].replace(0,'-')

                self.inserttext(str(Data_list.loc[i][3]))
                self.hwp.Run("TableRightCellAppend")
                try :
                    self.inserttext((format(Data_list.loc[i][4],'10.2E')))
                except :
                    self.inserttext((format(Data_list.loc[i][4])))

                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][5]))
                self.hwp.Run("TableRightCellAppend")
                try :
                    self.inserttext((format(Data_list.loc[i][6],'10.2E')))
                except :
                    self.inserttext((format(Data_list.loc[i][6])))

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

                self.inserttext(str(Data_list.loc[i][9]))
                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][10]))
                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][11]))
                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][12]))
                self.hwp.Run("TableRightCellAppend")
                self.inserttext(str(Data_list.loc[i][13]))



    def Main(self,host,port,database,user,password):

        dbname_f = self.DbName_f
        'hwp_4_4_f'
        dbname_p = self.DbName_p
        'hwp_4_4_p'


        db = dbupdate_update.DBCONN()
        db.SetDataBase(host,port,database,user,password)
        db.DB_CONN()

        hwp_4_3_f = db.query(dbname_f,"*")
        hwp_4_3_f = pd.DataFrame(hwp_4_3_f)

        hwp_4_3_p = db.query (dbname_p,"*")
        hwp_4_3_p = pd.DataFrame(hwp_4_3_p)


        gettype = self.GetProcessType(hwp_4_3_f,2)
        hwp_4_3_p.fillna('-',inplace=True)

        hwp_4_3_f[9] = hwp_4_3_f[9].str.replace(',',',\n')
        hwp_4_3_f[10] = hwp_4_3_f[10].str.replace(',',',\n')

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
            text = hwp_table[18][0].split("-")
            name = Manufa_type[text[0]]+"-"+text[1]+"-"+text[2]
            self.inserttext_notall(name)

            for j in range(0,int(len(hwp_table))):
        
        
                contrl_type_list, contrl_type, Select_data = self.searchTbl_p(hwp_4_3_p, hwp_table, j)

        
                self.DelTbl_p(contrl_type, contrl_type_list)

                type_count = self.MakeTableCol_p(contrl_type, contrl_type_list)


                self.InsertTbl_f(hwp_table, j)


                self.InsertTbl_p(Select_data,contrl_type,type_count)


                self.Copy_Update_table()
                self.Copy_Table_sample()
                
                
            name = name + '_4장_' + self.fileName
            
            self.hwp.XHwpDocuments.Item(2).SetActive_XHwpDocument()
            time.sleep(0.1)
            self.hwp.SaveAs(os.path.join(path, name +'.hwp'))
            time.sleep(0.1)
            self.hwp.Run("FileClose")
            time.sleep(0.1)
            self.hwp.Run("FileNew")
            self.hwp.XHwpWindows.Item(2).Visible = False
            time.sleep(0.1)



        self.hwp.XHwpWindows.Item(2).Visible = True
        self.hwp.Run("FileClose")
        self.hwp.XHwpWindows.Item(1).Visible = True
        self.hwp.SaveAs(os.path.join(path,'trash_4장배출시설.hwp'))
        time.sleep(0.1)
        self.hwp.Run("FileClose")
        self.hwp.XHwpWindows.Item(0).Visible = True
        time.sleep(0.1)
        self.hwp.Run("FileClose")
        time.sleep(0.1)


        self.hwp.Quit()
        
        
        
        
        

    













