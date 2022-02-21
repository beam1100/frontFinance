from macro_print import *

import tkinter as tk
from tkinter import Scrollbar, ttk
import tkinter.messagebox as msgbox
import tkinter.ttk as myCombo
from tkinter.constants import ANCHOR


import calendar
from datetime import datetime, date
# from datetime import date, datetime
from tkinter import filedialog
import pandas as pd

import copy
import requests
import json

today = date.today() #날짜 값만 필요
lastDay = calendar.monthrange(today.year, today.month)[1]

import sqlite3

con = sqlite3.connect('./db/myDb.db')
with con:
    cursor = con.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS my_list(id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, dict TEXT) ")
    con.commit()


import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import pickle





class macroFrame(tk.Frame):

    selected = []

    def __init__(self):

        super().__init__()

        #선택된 지표 트리뷰
        self.selectedTree = ttk.Treeview(self, column=('c0', 'c1'), show='headings', height=30)
        self.selectedTree.heading('c0', text='이름')
        self.selectedTree.column('c0', width=200, anchor=tk.CENTER)
        self.selectedTree.heading('c1', text='변화율?')
        self.selectedTree.column('c1', width=200, anchor=tk.CENTER)
        self.selectedTree.grid(row=0, column=0, rowspan=7)
        self.selectedTree.bind('<Button-3>', lambda event:self.rightClickEvent(event))


        # 이벤트 메뉴
        self.m = tk.Menu(self, tearoff = 0)
        self.m.add_command(label ='삭제', command=lambda:self.delFromSelected())
        self.m.add_command(label ='변화율 지정/해제', command=lambda:self.toChangeRate())
        self.m.add_command(label ='순서 변경', command=lambda:self.changePosition())


        # 시작일        
        yearList = [str(i)+'년' for i in range(1980, today.year+1)]
        self.startYearCombo = myCombo.Combobox(self, height=0, values=yearList, state='readonly')
        self.startYearCombo.current(len(yearList)-1)
        self.startYearCombo.grid(row=0, column=1, padx= 5)

        monthList = [str(i)+'월' for i in range(1, 13)]
        self.startMonthCombo = myCombo.Combobox(self, height=0, values=monthList, state='readonly')
        self.startMonthCombo.current(today.month-1)
        self.startMonthCombo.bind("<<ComboboxSelected>>", lambda event : self.selectDate(event, 'startMonth'))
        self.startMonthCombo.grid(row=0, column=3, padx= 5)

        self.startDayCombo = myCombo.Combobox(self, height=0)
        self.startDayValue = [str(i)+'일' for i in range(1, lastDay+1)]
        self.startDayCombo = myCombo.Combobox(self, height=0, values=self.startDayValue, state='readonly')
        self.startDayCombo.current(0)
        self.startDayCombo.grid(row=0, column=5, padx= 5)
        
        tk.Label(self, text='에서').grid(row=0, column=6, padx= 5, sticky=tk.W)


        # 종료일
        self.endYearCombo = myCombo.Combobox(self, height=0, values=yearList, state='readonly')
        self.endYearCombo.current(len(yearList)-1)
        self.endYearCombo.grid(row=1, column=1, padx= 5)

        self.endMonthCombo = myCombo.Combobox(self, height=0, values=monthList, state='readonly')
        self.endMonthCombo.current(today.month-1)
        self.endMonthCombo.bind("<<ComboboxSelected>>", lambda event:self.selectDate(event, 'endMonth'))
        self.endMonthCombo.grid(row=1, column=3, padx= 5)

        self.endDayValue = [str(i)+'일' for i in range(1, lastDay+1)]
        self.endDayCombo = myCombo.Combobox(self, height=0, values=self.endDayValue, state='readonly')
        self.endDayCombo.current(today.day-1)
        self.endDayCombo.grid(row=1, column=5, padx= 5)

        tk.Label(self, text='까지').grid(row=1, column=6, padx= 5, sticky=tk.W)


        # ecos 리스트박스
        tk.Label(self, text='한국은행에서').grid(row=2, column=1)

        self.ecosSelect = tk.Listbox(self, selectmode='single', height=20)
        self.ecosList = []
        con = sqlite3.connect('./db/myDb.db')
        with con:
            sqlite3.Connection
            cursor = con.cursor()
            sql = 'select * from ecos_indicator order by name'
            cursor.execute(sql)
            rows = cursor.fetchall()
            for row in rows:
                self.ecosSelect.insert(tk.END,  row[1])
                code2 = []
                for code in [row[4], row[5], row[6]]:
                    if code != '':
                        code2.append(code)
                self.ecosList.append({
                    'source':'ecos',
                    'name':row[1],
                    'code1':row[2],
                    'period':row[3],
                    'code2':code2
                })

        # ecos 스크롤바
        self.ecosScroll = Scrollbar(self, orient='vertical')
        self.ecosSelect.config(yscrollcommand=self.ecosScroll.set)
        self.ecosScroll.config(command=self.ecosSelect.yview)
        self.ecosScroll.grid(row=3, column=2, sticky=tk.N + tk.S + tk.W)

        #ecos 리스트박스 이벤트 연결 및 배치
        self.ecosSelect.bind('<Double-Button-1>', lambda event:self.doubleClickEvent(event, 'ecos'))
        self.ecosSelect.grid(row=3, column=1, sticky=tk.N + tk.E)


        # yf 리스트 박스
        tk.Label(self, text='야후파이낸스에서').grid(row=2, column=3)
        self.yfSelect = tk.Listbox(self, selectmode='single', height=20)
        self.yfList = []
        with con:
            sqlite3.Connection
            cursor = con.cursor()
            sql = 'select * from yf_indicator order by name'
            cursor.execute(sql)
            rows = cursor.fetchall()
            for row in rows:
                self.yfSelect.insert(tk.END,  row[1])
                self.yfList.append({
                    'source':'yf',
                    'name':row[1],
                    'symbol':row[2]
                })

        # yf 스크롤바
        self.yfScroll = Scrollbar(self, orient='vertical')
        self.yfSelect.config(yscrollcommand=self.yfScroll.set)
        self.yfScroll.config(command=self.yfSelect.yview)
        self.yfScroll.grid(row=3, column=4, sticky=tk.N + tk.S + tk.W)

        #yf 리스트박스 이벤트 연결 및 배치
        self.yfSelect.grid(row=3, column=3, sticky=tk.N + tk.E)
        self.yfSelect.bind('<Double-Button-1>', lambda event:self.doubleClickEvent(event, 'yf'))


        # fred 리스트 박스
        tk.Label(self, text='fred에서').grid(row=2, column=5)
        self.fredSelect = tk.Listbox(self, selectmode='single', height=20)
        self.fredList = []
        with con:
            sqlite3.Connection
            cursor = con.cursor()
            sql = 'select * from fred_indicator order by name'
            cursor.execute(sql)
            rows = cursor.fetchall()
            for row in rows:
                self.fredSelect.insert(tk.END,  row[1])
                self.fredList.append({
                    'source':'fred',
                    'name':row[1],
                    'code':row[2],
                    'period':row[3]
                })


        # fred 스크롤바
        self.fredScroll = Scrollbar(self, orient='vertical')
        self.fredSelect.config(yscrollcommand=self.fredScroll.set)
        self.fredScroll.config(command=self.fredSelect.yview)
        self.fredScroll.grid(row=3, column=6, sticky=tk.N + tk.S + tk.W)

        #fred 리스트박스 이벤트 연결 및 배치
        self.fredSelect.grid(row=3, column=5, sticky=tk.N)
        self.fredSelect.bind('<Double-Button-1>', lambda event:self.doubleClickEvent(event, 'fred'))


        # krx 개별종목 코드입력
        tk.Label(self, text='krx에서\n(종목코드입력)').grid(row=2, column=7)
        self.e= tk.Entry(self, width=20)
        self.e.grid(row=3, column=7, sticky=tk.N)
        self.e.bind('<Return>',  lambda event:self.doubleClickEvent(event, 'krx'))


        # # 내 목록 가져오기
        # btnSaveList = tk.Button(self, text='리스트 저장하기', command=lambda:self.myList('save'))
        # btnSaveList.grid(row=4, column=1)

        # btnLoadList  = tk.Button(self, text='리스트 가져오기', command=lambda:self.myList('load'))
        # btnLoadList.grid(row=4, column=3)
        
        # btnDelList  = tk.Button(self, text='리스트 삭제', command=lambda:self.myList('del'))
        # btnDelList.grid(row=4, column=5)

        # # btnClearList  = tk.Button(self, text='리스트 테이블 드랍', command=lambda:self.myList('clear'))
        # # btnClearList.grid(row=4, column=4)





        # 저장버튼
        btnRun = tk.Button(self, text='가져오기', command=lambda:self.getData())
        btnRun.grid(row=4, column=3, padx=1, pady=1)


        # 기준 주기 콤보
        self.spList = [
            {'name':'원자료 주기', 'sp':None}, 
            {'name':'QQ(분기)', 'sp':'QQ'}, 
            {'name':'MM(월)', 'sp':'MM'},
            {'name':'DD(일)', 'sp':'DD'},
        ]
        self.spCombo = myCombo.Combobox(self, height=0, width=20, values=[item['name'] for item in self.spList], state='readonly')
        self.spCombo.current(0)
        self.spCombo.grid(row=4, column=1, columnspan=2)





        # 불러오기 버튼
        btnRun = tk.Button(self, text='불러오기', command=lambda:self.loadData())
        btnRun.grid(row=5, column=3, padx=1, pady=1)

        self.picklePath = tk.Entry(self, width=20)
        self.picklePath.grid(row=5, column=1, columnspan=2)




        # 출력하기 버튼
        btnRun = tk.Button(self, text='출력', command=lambda:self.printData())
        btnRun.grid(row=6, column=3, padx=1, pady=1)


        # 출력기능 콤보
        self.functionList = [
            {'name':'그래프출력', 'param':'graph'},
            {'name':'엑셀파일로 출력', 'param':'cor'},
            {'name':'수익률 비교', 'param':'ccr'},   # Comparison of Cumulative Returns
            {'name':'히스토그램(누적수익률)', 'param':'histo'},
            {'name':'이동평균분석', 'param':'ma'},
            {'name':'볼린저밴드', 'param':'band'},
        ]
        self.functioinCombo = myCombo.Combobox(self, height=0, width=20, values=[item['name'] for item in self.functionList], state='readonly')
        self.functioinCombo.current(0)
        self.functioinCombo.grid(row=6, column=1, columnspan=2)




















    def getData(self):

        # print(macroFrame.selected)

        # 기준 주기 선택
        periodList = [] 
        for dict in macroFrame.selected:
            if 'period' in dict.keys():
                periodList.append(dict['period'])
        for item in self.spList:
            if item['name'] == self.spCombo.get():
                selectedSp = item['sp']
        if selectedSp:
            sp = selectedSp
        else:
            if 'QQ' in periodList:
                sp = 'QQ'       #sp: standart Period
            elif 'MM' in periodList:
                sp = 'MM'
            else:
                sp = 'DD'

        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        # path = os.getcwd()   #현재 폴더 경로; 작업 폴더 기준
        # fileList = os.listdir(os.getcwd())    #디렉토리 파일 리스트
        filePath = os.path.dirname(os.path.realpath(__file__))     #현재 파일의 폴더 경로; 작업 파일 기준


        #다운로드 경로 지정
        options.add_experimental_option(
            'prefs',
            {
                'download.default_directory' : filePath + '\\pickles'
            }
        ) 

        browser = webdriver.Chrome('./chromedriver', options=options)
        browser.get('http://114.199.29.251:8000/myinput')
        startDay, endDay = self.getDate()

        browser.find_element(By.NAME, 'startDay').send_keys(startDay)
        browser.find_element(By.NAME, 'endDay').send_keys(endDay)
        browser.find_element(By.NAME, 'sp').send_keys(sp)
        dictList = json.dumps(macroFrame.selected)
        browser.find_element(By.NAME, 'dictList').send_keys(dictList)
        browser.find_element(By.ID, 'submit').click()

        time.sleep(5)
        browser.find_element(By.ID, 'down').click()
        time.sleep(5)

        browser.quit()



















    def loadData(self):
        value = filedialog.askopenfile(parent=self, initialdir='./pickles', title='저장된 피클 선택', filetypes=[('pickle', '.*')])
        if value == None:
            return
        else:
            self.picklePath.delete(0, tk.END)
            self.picklePath.insert(0, value.name)



















    def printData(self):
        _dfSum = open(self.picklePath.get(), 'rb')
        dfSum = pickle.load(_dfSum)
        
        # # 기능 선택
        for item in self.functionList:
            if item['name'] == self.functioinCombo.get():
                selectedCombo = item['param']
                break

        #엑셀파일로 저장되는 기능은 파일경로 요청
        dir = None
        if selectedCombo == 'cor':
            dir = filedialog.asksaveasfile(mode='w', defaultextension=".xlsx")
            if dir == None:
                return

        # print(dfSum)

        if selectedCombo == 'graph':
            createGraph(dfSum, 'nomal')

        elif selectedCombo == 'ccr':
            createGraph(dfSum, 'ccr')

        elif selectedCombo == 'histo':
            createGraph(dfSum, 'histogram')

        elif selectedCombo == 'ma':
            createGraph(dfSum, 'ma')

        elif selectedCombo == 'band':
            createGraph(dfSum, 'band')

        elif selectedCombo == 'cor':
            createExcel(dfSum, dir.name, 'cor')



















    def getDate(self):
        startYear = int(self.startYearCombo.get().replace('년',''))
        startMonth = int(self.startMonthCombo.get().replace('월',''))
        startDay = int(self.startDayCombo.get().replace('일',''))
        endYear = int(self.endYearCombo.get().replace('년',''))
        endMonth = int(self.endMonthCombo.get().replace('월',''))
        endDay = int(self.endDayCombo.get().replace('일',''))
        start = date(startYear, startMonth, startDay).strftime('%Y%m%d')
        end = date(endYear, endMonth, endDay).strftime('%Y%m%d')
        return start, end





    def doubleClickEvent(self, event, widget):
        if widget == 'ecos':
            index = self.ecosSelect.curselection()[0]
            dict = copy.deepcopy(self.ecosList[index]) #얕은 복사는 참조하기 때문에, 깊은 복사 필요
            dict['changeRate'] = False
        elif widget =='yf':
            index = self.yfSelect.curselection()[0]
            dict = copy.deepcopy(self.yfList[index])
            dict['changeRate'] = False
        elif widget == 'fred':
            index = self.fredSelect.curselection()[0]
            dict = copy.deepcopy(self.fredList[index])
            dict['changeRate'] = False
        elif widget == 'krx':
            code = self.e.get()
            # name = stock.get_market_ticker_name(code)
            dict = {'source':'krx', 'name':code, 'code':code, 'changeRate':False}
        
        macroFrame.selected.append(dict)

        self.updateTree()





    def delFromSelected(self):
        items = self.selectedTree.selection()
        itemList = self.selectedTree.get_children()
        
        for item in reversed(items):
            index = itemList.index(item)
            del macroFrame.selected[index]
        self.updateTree()





    def toChangeRate(self):
        item = self.selectedTree.focus()
        itemList = self.selectedTree.get_children()
        index = itemList.index(item)
        macroFrame.selected[index]['changeRate'] = not(macroFrame.selected[index]['changeRate'])
        self.updateTree()





    def changePosition(self):
        items = self.selectedTree.selection()
        itemList = self.selectedTree.get_children()
        if len(items)== 2:
            index0 = itemList.index(items[0])
            index1 = itemList.index(items[1])
            macroFrame.selected[index0], macroFrame.selected[index1] = macroFrame.selected[index1], macroFrame.selected[index0]
        else:
            msgbox.showwarning('경고','두 개를 선택하세요')
        self.updateTree()





    def selectDate(self, event, widget):
        startYear = int(self.startYearCombo.get().replace('년',''))
        startMonth = int(self.startMonthCombo.get().replace('월',''))
        startDay = int(self.startDayCombo.get().replace('일',''))
        endYear = int(self.endYearCombo.get().replace('년',''))
        endMonth = int(self.endMonthCombo.get().replace('월',''))
        endDay = int(self.endDayCombo.get().replace('일',''))
        start = date(startYear, startMonth, startDay)
        end = date(endYear, endMonth, endDay)
        startEndDiff = (end - start).days
        todayEndDiff = (today - end).days
        
        if widget == 'startMonth':
            lastDay = calendar.monthrange(startYear, startMonth)[1]
            _startDayList = list(range(1,lastDay+1))
            startDayList = [str(day)+'일' for day in _startDayList]
            self.startDayCombo.configure(values=startDayList)
            self.startDayCombo.current(0)
        elif widget == 'endMonth':
            lastDay = calendar.monthrange(endYear, endMonth)[1]
            _endDayList = list(range(1,lastDay+1))
            endDayList = [str(day)+'일' for day in _endDayList]
            self.endDayCombo.configure(values=endDayList)
            self.endDayCombo.current(0)

        if startEndDiff < 0:
            msgbox.showwarning('경고','종료일이 시작일보다 과거입니다. 날짜를 다시 지정하세요')
        
        if todayEndDiff < 0:
            msgbox.showwarning('경고','종료일이 미래입니다. 날짜를 다시 지정하세요')





    # def myList(self, param):
        
    #     with con:
    #         cursor = con.cursor()
    #         if param == 'save':
    #             cursor.execute('delete from my_list')
    #             con.commit()
    #             for item in macroFrame.selected:
    #                 sql = 'INSERT INTO my_list (dict) values(?)'
    #                 _item = json.dumps(item)
    #                 cursor.execute( sql, [(_item)])
    #                 con.commit()

    #         elif param == 'load':
    #             sql = 'SELECT * FROM my_list'
    #             cursor.execute(sql)
    #             results = cursor.fetchall()
    #             selectedLoaded = []
    #             for result in results:
    #                 selectedLoaded.append(json.loads(result[1]))
    #             macroFrame.selected = selectedLoaded
    #             self.updateTree()
                    
    #         elif param == 'del':
    #             sql = 'delete from my_list'
    #             cursor.execute(sql)
    #             con.commit()
    #             print('my_list 모든 row 삭제됨')

    #         elif param == 'clear':
    #             cursor.execute('DROP TABLE my_list')
    #             con.commit()
    #             print('my_list 테이블 드롭됨')

                    



    def rightClickEvent(self, event):
        self.m.tk_popup(event.x_root, event.y_root)
        self.m.grab_release()





    def updateTree(self):
        #기존에 있던 목록 삭제
        for item in self.selectedTree.get_children():
            self.selectedTree.delete(item)
        #새로운 목록 입력
        for dic in macroFrame.selected:
            if dic['changeRate']:
                self.selectedTree.insert(
                    '',
                    'end',
                    text='',
                    values=(dic['name'], 'O')
                )
            else:
                self.selectedTree.insert(
                    '',
                    'end',
                    text='',
                    values=(dic['name'], 'X')
                )










# 테스트 바로 실행
if __name__ == '__main__':
    root = tk.Tk()
    macroFrame().pack()
    root.mainloop()
