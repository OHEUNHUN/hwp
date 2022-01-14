
# -*- coding: utf-8 -*-
"""
Created on Tue Oct 12 16:13:49 2021

한글 입력 프로그램 v1.01버젼

업데이트 사항-->

1) 1장, 4장의 Set_propertyㅇ에 대한 값 오류를 수정함.
2) 2장, 3장에 대한 출력 프로그램 추가
3) DB접속시도를 무제한으로 실행가능 --> 이후 json파일을 이용한 방식으로 수정예정

@author: ECOCNA_dev
"""
import chapter_1_방지시설
import chapter_1_배출시설
import chapter_1_비대상방지시설
import chapter_1_비대상배출시설
import chapter_2_최대배출기준
import chapter_2_예상배출기준
import chapter_2_허가배출기준
import chapter_3_허가배출기준
import chapter_4_방지시설
import chapter_4_배출시설
import dbupdate_update as db
import mariadb


class Hwp_exe:
    
    def __init__(self):
        
        super().__init__()
        self.chapter_1_1 = chapter_1_배출시설.chapter_1_production()
        self.chapter_1_2 = chapter_1_방지시설.chapter_1_protect()
        self.chapter_1_1_1 = chapter_1_비대상배출시설.chapter_1_production()
        self.chapter_1_2_1 = chapter_1_비대상방지시설.chapter_1_protect()
        self.chapter_2_1_1 = chapter_2_최대배출기준.chapter_2_Max()
        self.chapter_2_1_2 = chapter_2_예상배출기준.chapter_2_Predic()
        self.chapter_2_2 = chapter_2_허가배출기준.chapter_2_Premission()
        self.chapter_3_1 = chapter_3_허가배출기준.chapter_3_Premission()
        self.chapter_4_1 = chapter_4_배출시설.chapter_4_production()
        self.chapter_4_2 = chapter_4_방지시설.chapter_4_protect()



    def Printproperty(self,Data ,Num):
        
        if Num == 1:
            print('샘플경로: {} \n샘플 이름: {} \nDB명: {} \n파일경로: {} \n파일이름: {} \n'.format(Data[0],Data[1],Data[2],Data[3],Data[4]))
        
        if Num == 2:
            print('샘플경로: {} \n샘플 이름: {} \nDB명_f: {} \nDB명_p: {} \n파일경로: {} \n파일이름: {} \n'.format(Data[0],Data[1],Data[2],Data[3],Data[4],Data[5]))



    def Main(self):
        
        while True :
            host = input('host를 입력하세요.\n')
            port = input('port를 입력하세요.\n')
            database = input('database를 입력하세요.\n')
            user = input('user를 입력하세요.\n')
            password = input('password를 입력하세요.\n')

            db_check = db.DBCONN()
            db_check.SetDataBase(host,port,database,user,password)
            
            try:
                if(db_check.DB_CONN() == 'success'):
                    db_check.conn.close()
                    print("연결 성공")
                    break

            except mariadb.OperationalError:
                print('데이터베이스 입력 정보가 틀립니다.')
                input("아무거나 누르세요")
    
            except mariadb.ProgrammingError:
                print('데이터베이스 입력 정보가 틀립니다(한글입력).')
                input("아무거나 누르세요")
        
        while True:
            
            print(' 전체 출력 : 0\n 1장 출력  : 1\n 2장 출력  : 2\n 3장 출력  : 3\n 4장 출력  : 4\n 경로 출력 : path\n 종료 : exit\n')
            PrintCase = input("출력할 대상을 입력하세요\n")
        
        
            if PrintCase == "0" :
            
                print("전체 출력\n")

                self.chapter_1_1.Main(host,port,database,user,password)
                print("1장 배출 시설 출력 완료\n")
                self.chapter_1_2.Main(host,port,database,user,password)
                print("1장 방지 시설 출력 완료\n")
                self.chapter_1_1_1.Main(host,port,database,user,password)
                print("1장 비대상 배출 시설 출력 완료\n")
                self.chapter_1_2_1.Main(host,port,database,user,password)
                print("1장 비대상 방지 시설 출력 완료\n")
                self.chapter_2_1_1.Main(host,port,database,user,password)
                print("2장 최대배출기준 출력 완료\n")
                self.chapter_2_1_2.Main(host,port,database,user,password)
                print("2장 예상배출기준 출력 완료\n")
                self.chapter_2_2.Main(host,port,database,user,password)
                print("2장 허가배출기준 출력 완료\n")
                self.chapter_3_1.Main(host,port,database,user,password)
                print("3장 허가배출기준 출력 완료\n")
                self.chapter_4_1.Main(host,port,database,user,password)
                print("4장 배출 시설 출력 완료\n")
                self.chapter_4_2.Main(host,port,database,user,password)
                print("4장 방지 시설 출력 완료\n")

            elif PrintCase == "1" :
            
                print("1장 출력\n")

                self.chapter_1_1.Main(host,port,database,user,password)
                print("1장 배출 시설 출력 완료\n")
                self.chapter_1_2.Main(host,port,database,user,password)
                print("1장 방지 시설 출력 완료\n")
                self.chapter_1_1_1.Main(host,port,database,user,password)
                print("1장 비대상 배출 시설 출력 완료\n")
                self.chapter_1_2_1.Main(host,port,database,user,password)
                print("1장 비대상 방지 시설 출력 완료\n")


            elif PrintCase == "2":

            
                self.chapter_2_1_1.Main(host,port,database,user,password)
                print("2장 최대배출기준 출력 완료\n")

                self.chapter_2_1_2.Main(host,port,database,user,password)
                print("2장 예상배출기준 출력 완료\n")

                self.chapter_2_2.Main(host,port,database,user,password)
                print("2장 허가배출기준 출력 완료\n")



            elif PrintCase == "3":
                self.chapter_3_1.Main(host,port,database,user,password)
                print("3장 허가배출기준 출력 완료\n")


        
            elif PrintCase == "4":
            
                print("4장 출력\n")
                self.chapter_4_1.Main(host,port,database,user,password)
                print("4장 배출 시설 출력 완료\n")
                self.chapter_4_2.Main(host,port,database,user,password)
                print("4장 방지 시설 출력 완료\n")

            elif PrintCase == "path":

                print("1장 배출 시설\n")
                self.Printproperty(self.chapter_1_1.Get_property(), 1)
                print("1장 방지 시설\n")
                self.Printproperty(self.chapter_1_2.Get_property(), 1)

                print("1장 비대상배출 시설\n")
                self.Printproperty(self.chapter_1_1_1.Get_property(), 1)
                print("1장 비대상방지 시설\n")
                self.Printproperty(self.chapter_1_2_1.Get_property(), 1)


                print("2장 최대배출기준\n")
                self.Printproperty(self.chapter_2_1_1.Get_property(), 1)

                print("2장 최대배출기준\n")
                self.Printproperty(self.chapter_2_1_2.Get_property(), 1)

                print("2장 허가배출기준\n")
                self.Printproperty(self.chapter_2_2.Get_property(), 1)


                print("3장 허가배출기준\n")
                self.Printproperty(self.chapter_3_1.Get_property(), 1)


                print("4장 배출 시설\n")
                self.Printproperty(self.chapter_4_1.Get_property(), 2)
                print("4장 방지 시설\n")
                self.Printproperty(self.chapter_4_2.Get_property(), 2)

            elif PrintCase == 'exit':

                quit()

            else :
                print("해당 조건 없음")



if __name__ == "__main__":
    Hwp_exe().Main()
