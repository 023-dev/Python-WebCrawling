# -*- coding: utf-8 -*-
#Python : 3.6.8
#Date : 2020.02.06
#==============================================================================================================모듈
import os
import csv
import time
import requests
from tkinter import *
from tkinter import ttk
from bs4 import BeautifulSoup
from datetime import datetime
from selenium import webdriver
from urllib.request import urlopen
from urllib.parse import quote_plus
from tkinter import messagebox as msg
#==============================================================================================================디렉토리 경로 정의
if os.path.exists('C:\Issue') == True:
    os.chdir('C:\Issue')
elif os.path.exists('C:\Issue') == False:
    os.makedirs('C:\Issue')
    os.chdir('C:\Issue')
Csv_Path = 'C:\Issue'
#==============================================================================================================메시지, 실행 기능
def Execute_Check_Msg():
    msg.askokcancel("","실행전에 검색단어를 확인하세요.\n실행하시겠습니까?")
    if 'yes':
        Csv_Produce()
def Finish_Msg():
    msg.showinfo("","작업을 완료하였습니다.")
def Finish_Write_Msg():
    msg.showinfo("", "파일은 "+Csv_Path+ "에 저장되었습니다.")
def Customer_Append_Msg():
    msg.showinfo("","고객사를 추가하였습니다.")
def Customer_Append_Function():
    for Customer_Plus in Customer_Entry.get().split(","):
        Customer_List.append(Customer_Plus)
    Customer_Whole = int(len(Customer_List))
    Customer_Append_Msg()
def Customer_Delete_Msg():
    msg.showinfo("","고객사를 제거하였습니다.") 
def Customer_Delete_Function():
    for Customer_Minus in Customer_Entry.get().split(","):
        if Customer_Minus not in Customer_List:
            print("x not in list")        
        Customer_List.remove(str(Customer_Minus))
    Customer_Whole = int(len(Customer_List))
    Customer_Delete_Msg()
def Customer_Reset_Msg():
    msg.showinfo("","모든 고객사를 제거하였습니다.")
def Customer_Reset_Function():
    Customer_List.clear()
    Customer_Whole = int(len(Customer_List))
    Customer_Reset_Msg()
def Word_Append_Msg():
    msg.showinfo("","키워드를 추가하였습니다.") 
def Word_Append_Function():
    for Word_Plus in Word_Entry.get().split(","):
        Word_List.append(str(Word_Plus))   
    Word_Whole = int(len(Word_List))
    Word_Append_Msg()
def Word_Delete_Msg():
    msg.showinfo("","키워드를 제거하였습니다.")
def  Word_Delete_Function():
    for Word_Minus in Word_Entry.get().split(","):
        Word_List.remove(str(Word_Minus))  
    Word_Whole  = int(len(Word_List))
    Word_Delete_Msg()
def Word_Reset_Msg():
    msg.showinfo("","모든 키워드를 제거하였습니다.")
def  Word_Reset_Function():
    Word_List.clear()
    Word_Whole  = int(len(Word_List))
    Word_Reset_Msg()
#==============================================================================================================메인 프레임
Main_Frame =  Tk()
Main_Frame.title("")
Main_Frame.geometry("1350x760+0+0")
Main_Frame.resizable(0,0)
Main_Frame_Title = Label(Main_Frame, text = "고객사 이슈사항", bd = 6, relief = RIDGE, font = ("times new roman", 30 , "bold"), bg = "gray", fg = "white")
Main_Frame_Title.pack(side = TOP, fill = X)
#==============================================================================================================검색 조건 관련 프레임
Search_Frame = Frame(Main_Frame, bd = 6, relief = RIDGE, bg = "gray")
Search_Frame.place(x = 25, y =90, width = 450, height = 340)
Search_Term = Label(Search_Frame, text="검색할 기간을 입력하세요.", font = ("times new roman", 20 , "bold"), bg = "gray", fg = "white")
Search_Term.grid(row = 0, column =0, pady =2, padx =20, sticky ="w")
Search_Term_Example = Label(Search_Frame, text="예)2020.01.01", font = ("times new roman", 20 , "bold"), bg = "gray", fg = "white")
Search_Term_Example.grid(row = 1, column =0, pady =2, padx =20, sticky ="w")
From_term_Entry = Entry(Search_Frame, font = ("times new roman", 15 , "bold"), bd = 6, relief = RIDGE)
From_term_Entry.grid(row = 2, column =0, pady =2, padx =20, sticky ="w")
From_term_Label = Label(Search_Frame, text="부터", font = ("times new roman", 20 , "bold"), bg = "gray", fg = "white")
From_term_Label.grid(row = 2, column =0, pady =2, padx =240, sticky ="w")
To_term_Entry = Entry(Search_Frame, font = ("times new roman", 15 , "bold"), bd = 6, relief = RIDGE)
To_term_Entry.grid(row = 3, column =0, pady =2, padx =20, sticky ="w")
To_term_Label = Label(Search_Frame, text="까지", font = ("times new roman", 20 , "bold"), bg = "gray", fg = "white")
To_term_Label.grid(row = 3, column =0, pady =10, padx = 240, sticky ="w")
Page_Label = Label(Search_Frame, text="페이지 수를 입력하세요.", font = ("times new roman", 20 , "bold"), bg = "gray", fg = "white")
Page_Label.grid(row = 4, column =0, pady =2, padx =20, sticky ="w")
Page_Consider = Label(Search_Frame, text="*많은 페이지를 입력하면 정확성이 떨어집니다.(1페이지당 기사10개)", font = ("times new roman", 7 , "bold"), bg = "gray", fg = "white")
Page_Consider.grid(row = 5, column =0, pady =2, padx = 20, sticky ="w")
Page_Entry = Entry(Search_Frame, font = ("times new roman", 15 , "bold"), bd = 6, relief = RIDGE)
Page_Entry.grid(row = 6, column =0, pady =2, padx =20, sticky ="w")
#==============================================================================================================
#알터프레임
Alter_Frame = Frame(Main_Frame, bd = 6, relief = RIDGE, bg = "gray")
Alter_Frame.place(x = 25, y =430, width = 450, height = 300)
#==============================================================================================================
#고객사입력
Customer_Label = Label(Alter_Frame, text="추가 또는 제거할 회사를 입력하세요.", font = ("times new roman", 16 , "bold"), bg = "gray", fg = "white")
Customer_Label.grid(row = 0, column =0, pady =2, padx =20, sticky ="w")
Customer_Entry = Entry(Alter_Frame, font = ("times new roman", 15 , "bold"), bd = 6, relief = RIDGE)
Customer_Entry.grid(row = 1, column =0, pady =2, padx =20, sticky ="w")
Customer_Append_Button = Button(Alter_Frame, text = "추가",  bg = "light grey", width =10, bd = 6, relief = RAISED, command=Customer_Append_Function)
Customer_Append_Button.grid(row = 1, column =0, padx = 240, sticky ="w")
Customer_Delete_Button = Button(Alter_Frame, text = "제거",  bg = "light grey", width =10, bd = 6, relief = RAISED, command=Customer_Delete_Function)
Customer_Delete_Button.grid(row = 1, column =0, padx = 340, sticky ="w")
Customer_Reset_Button = Button(Alter_Frame, text = "모든 고객사 제거하기",  bg = "light grey", width =24, bd = 6, relief = RAISED, command=Customer_Reset_Function)
Customer_Reset_Button.grid(row = 2, column =0, padx = 240, sticky ="w")
#단어입력
Word_Label = Label(Alter_Frame, text="추가 또는 삭제할 검색 단어를 입력하세요.", font = ("times new roman", 16 , "bold"), bg = "gray", fg = "white")
Word_Label.grid(row = 3, column =0, pady =2, padx =20, sticky ="w")
Word_Entry = Entry(Alter_Frame, font = ("times new roman", 15 , "bold"), bd = 6, relief = RIDGE)
Word_Entry.grid(row = 4, column =0, pady =2, padx =20, sticky ="w")
Word_Append_Button = Button(Alter_Frame, text = "추가",  bg = "light grey", width =10, bd = 6, relief = RAISED, command=Word_Append_Function)
Word_Append_Button.grid(row = 4, column =0, padx = 240, sticky ="w")
Word_Delete_Button = Button(Alter_Frame, text = "제거",  bg = "light grey", width =10, bd = 6, relief = RAISED, command=Word_Delete_Function)
Word_Delete_Button.grid(row = 4, column =0, padx = 340, sticky ="w")
Word_Reset_Button = Button(Alter_Frame, text = "모든 검색 단어를  제거하기",  bg = "light grey", width =24, bd = 6, relief = RAISED, command=Word_Reset_Function)
Word_Reset_Button.grid(row = 5, column =0, padx = 240, sticky ="w")
#==============================================================================================================
#리스트
global Customer_List
global Word_List
global Result_List
Customer_List = ['LG화학', 'SK바이오랜드', 'GC녹십자', '다림티센', '대웅제약', '대원제약', '대한약품공업','더마펌', '더스탠다드', '덱스레보', '동국제약','동방메디컬', '루먼바이오', '메디파크', '메디톡스', '메타바이오메드', '바이오플러스', '비씨월드제약', '미알팜', '삼일제약', '삼천당제약', '세원셀론텍', '센트럴메디컬서비스', '셀트리온', '시지바이오', '아리바이오', '에스테팜', '에이치피앤씨', '엑소코바이오', '유바이오로직스', '유바이오메드', '이연제약', '제네웰', '제노스', '제테마', '종근당바이오', '중헌제약', '케어젠', '콘텍코리아', '파마리서치프로덕트', '펜믹스', '풍림파마텍', '한국비엔씨', '한국유나이티드제약', '한국코러스제약', '한국팜비오', '한미약품', '현대메디텍', '휴온스', '루멘스', '인벤티지랩', '와이솔', 'LG이노텍', '폴라이브', '엘씨스퀘어']
Word_List = ['식약처', '의약품', '제약', '약사법', '약', '임상']
Result_List = []
#==============================================================================================================
#검색어검토기능
def Check_Function():
    Customer_Whole = int(len(Customer_List))
    Word_Whole  = int(len(Word_List))
    msg.showinfo("", "고객사 : " + str(Customer_List) + "\n" + "관련 검색어 :  "+str(Word_List) )
#검색어검토버튼
Check_Function_Label = Label(Alter_Frame, text="*검색전에 검색 단어들을 확인하세요.", font = ("times new roman", 10 , "bold"), bg = "gray", fg = "white")
Check_Function_Label.grid(row = 6, column =0, pady =2, padx = 20, sticky ="w")
Check_Function_Button = Button(Alter_Frame, text = "검색단어 확인하기",  bg = "light grey", width =24, bd = 6, relief = RAISED, command=Check_Function)
Check_Function_Button.grid(row = 7, column =0, pady = 5,padx = 240, sticky ="w")
#==============================================================================================================크롤링 기능 코드
#CSV
def Csv_Produce():
    global Result_List
    global Customer_List
    global Word_List
    Customer_Whole = int(len(Customer_List))
    Word_Whole  = int(len(Word_List))
    From_term = From_term_Entry.get()
    To_term = To_term_Entry.get()
    To_Page = (int(Page_Entry.get())-1)*10+1 
    Duplicate_List = []
    Csv_Form = ['고객사', '이슈사항', '기사일자', '관련 기사 링크']
    Result_List.append(Csv_Form)
    Page = 1
    Customer_Increase = 0
    Word_Increase = 0
    while Customer_Increase < Customer_Whole:
        while Word_Increase < Word_Whole:
            '''print(Link_Append)'''
            while Page <= To_Page:
                if int(Page_Entry.get()) == 1:
                    Link = "https://search.naver.com/search.naver?where=news&query="+quote_plus(str(Customer_List[Customer_Increase]))+"%20%2B"+quote_plus(str([Word_List[Word_Increase]]).replace('[','').replace(']','').replace("'",''))+"&sm=tab_opt&sort=0&photo=0&field=0&reporter_article=&pd=3&ds="+quote_plus(str(From_term))+"&de="+quote_plus(str(To_term))+"&docid=&nso=so%3Ar%2Cp%3Afrom"+quote_plus(str(From_term).replace('.',''))+"to"+quote_plus(str(To_term).replace('.',''))+"%2Ca%3Aall&mynews=0&refresh_start=0&related=0"
                else:
                    Link = "https://search.naver.com/search.naver?where=news&query="+quote_plus(str(Customer_List[Customer_Increase]))+"%20%2B"+quote_plus(str([Word_List[Word_Increase]]).replace('[','').replace(']','').replace("'",''))+"&sm=tab_opt&sort=0&photo=0&field=0&reporter_article=&pd=3&ds="+quote_plus(str(From_term))+"&de="+quote_plus(str(To_term))+"&docid=&nso=so%3Ar%2Cp%3Afrom"+quote_plus(str(From_term).replace('.',''))+"to"+quote_plus(str(To_term).replace('.',''))+"%2Ca%3Aall&mynews=0&start="+quote_plus(str(Page))+"11&refresh_start=0"
                print(Link)
                Link_Open = urlopen(Link).read()
                Link_Analyze = BeautifulSoup(Link_Open, 'html.parser')
                for Data in Link_Analyze.select(".news_tit"):
                    print(Data)
                    News = []
                    News.clear()
                    if Data.attrs['title'] not in Duplicate_List:
                        News.append(Customer_List[Customer_Increase])
                        News.append(Data.attrs['title'])
                        Duplicate_List.append(Data.attrs['title'])
                        News.append(str(From_term)+"~"+str(To_term))
                        News.append(Data.attrs['href'])
                        Result_List.append(News)     
                        print(Result_List)       
                Page += 10
            Word_Increase += 1
            Page = 1
        Customer_Increase += 1
        Word_Increase = 0
    Finish_Msg()
#==============================================================================================================csv생성 기능
def Csv_Writer():
    Date = time.strftime('%Y-%m-%d-%H')
    Csv_Setting = open(str(Date)+'고객사이슈사항.csv', 'w', encoding='utf-8' , newline = '')
    Csv_Write = csv.writer(Csv_Setting)
    for Result_Date in Result_List:    
        Csv_Write.writerow(Result_Date)
    Csv_Setting. close()
    Finish_Write_Msg()
#==============================================================================================================실행 버튼
#실행버튼
Execute_Frame = Frame(Search_Frame, bd = 6, relief = RIDGE, bg = "gray")
Execute_Frame.place(x=15,y = 287, width = 420)
Execute_Button = Button(Execute_Frame, text = "실행",  bg = "light grey", width =10, relief = RAISED, command = Execute_Check_Msg).pack(fill = X)
#==============================================================================================================고객사 조회
def Result_Inquiry_Nan():
    List_Inquiry = Inquiry_Entry.get()
    msg.showinfo("",List_Inquiry+"에 대한 기사를 찾지 못했습니다.")
def Result_Inquiry():
    global List_Inquiry
    List_Inquiry = Inquiry_Entry.get()
    for Tree_Data in Result_List:
        if List_Inquiry == Tree_Data[0]:
            Tree.insert('', 'end',values=(Tree_Data[0], Tree_Data[1], Tree_Data[2], Tree_Data[3]))
        else:
            Result_Inquiry_Nan
#===============================================================================트리 리셋 기능
def Tree_Reset():
    Tree.delete(*Tree.get_children())
#===============================================================================드라이버 정의
def Web_Driver():
    Web_Drive = webdriver.Chrome("C:\chromedriver")
    Web_Drive_Link=Tree.item(Tree.selection())['values'][3]
    Web_Drive.get(Web_Drive_Link)
#===============================================================================세부 기능 버튼
Sub_Frame = Frame(Main_Frame, bd = 6, relief = RIDGE, bg = "gray")
Sub_Frame.place(x = 500, y =90, width = 820, height = 640)
Inquiry_Entry_Consider = Label(Sub_Frame, text="*조회할 고객사명을 입력하세요.", font = ("times new roman", 10 , "bold"), bg = "gray", fg = "white")
Inquiry_Entry_Consider.grid(row = 0, column =1, pady =2, padx = 20, sticky ="w")
Inquiry_Entry = Entry(Sub_Frame, font = ("times new roman", 15 , "bold"), bd = 6, relief = RIDGE)
Inquiry_Entry.grid(row = 1, column =1, padx =20,pady =2, sticky ="w")
List_Inquiry = Inquiry_Entry.get()
Inquiry_Button = Button(Sub_Frame, text = "조회", bg = "light grey", width =10, bd = 6, relief = RAISED, command=Result_Inquiry)
Inquiry_Button.grid(row = 0, column =2,padx = 10,pady =5, sticky ="w")
Web_Driver_Button = Button(Sub_Frame, text = "링크열기", bg = "light grey", width =10, bd = 6, relief = RAISED, command=Web_Driver)
Web_Driver_Button.grid(row = 0, column =2,padx = 110,pady =5, sticky ="w")
Tree_Reset_Button = Button(Sub_Frame, text = "목록 지우기", bg = "light grey", width =10, bd = 6, relief = RAISED, command=Tree_Reset)
Tree_Reset_Button.grid(row = 0, column =2,padx = 210,pady =5, sticky ="w")
Csv_Writer_Button = Button(Sub_Frame, text = "CSV파일로 만들기", bg = "light grey", width =40, bd = 6, relief = RAISED, command=Csv_Writer)
Csv_Writer_Button.grid(row = 1, column =2,padx = 10,pady =2, sticky ="w")
#===============================================================================트리뷰 프레임
Tree_Frame = Frame(Sub_Frame, bd = 6, relief = RIDGE, bg = "gray")
Tree_Frame.place(x = 15, y =90, width = 780, height = 500)
Scroll_Bar_Y = Scrollbar(Tree_Frame, orient=VERTICAL)
Scroll_Bar_X = Scrollbar(Tree_Frame, orient=HORIZONTAL)
Tree = ttk.Treeview(Tree_Frame, columns=("고객사", "이슈사항", "기사일자","주소"), selectmode="extended", height=300,  yscrollcommand=Scroll_Bar_Y.set, xscrollcommand=Scroll_Bar_X.set)
Scroll_Bar_Y.config(command=Tree.yview)
Scroll_Bar_Y.pack(side=RIGHT, fill=Y)
Scroll_Bar_X.config(command=Tree.xview)
Scroll_Bar_X.pack(side=BOTTOM, fill=X)
Tree.heading('고객사', text="고객사", anchor=W)
Tree.heading('이슈사항', text="이슈사항", anchor=W)
Tree.heading('기사일자', text="기사일자", anchor=W)
Tree.heading('주소', text="주소", anchor=W)        
Tree.column('#0', stretch=NO, minwidth=0, width=0)
Tree.column('#1', stretch=NO, minwidth=0, width=100)
Tree.column('#2', stretch=NO, minwidth=0, width=320)
Tree.column('#3', stretch=NO, minwidth=0, width=140)
Tree.column('#4', stretch=NO, minwidth=0, width=200)
Tree.pack()
#===============================================================================
Main_Frame = mainloop()
