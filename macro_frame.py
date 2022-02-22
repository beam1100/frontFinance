from macro_print import *

import tkinter as tk
from tkinter import Checkbutton, Scrollbar, ttk
import tkinter.messagebox as msgbox
import tkinter.ttk as tkCombo
from tkinter.constants import ANCHOR

import calendar
from datetime import date

from tkinter import filedialog

import copy
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

################################################################################################################################################################################################################

        #선택된 지표 트리뷰
        self.selectedTree = ttk.Treeview(self, column=('c0', 'c1'), show='headings', height=30)
        self.selectedTree.heading('c0', text='이름')
        self.selectedTree.column('c0', width=200, anchor=tk.CENTER)
        self.selectedTree.heading('c1', text='변화율?')
        self.selectedTree.column('c1', width=200, anchor=tk.CENTER)
        self.selectedTree.grid(row=0, column=0, rowspan=3)
        self.selectedTree.bind('<Button-3>', lambda event:self.rightClickEvent(event))


        # 이벤트 메뉴
        self.m = tk.Menu(self, tearoff = 0)
        self.m.add_command(label ='삭제', command=lambda:self.delFromSelected())
        self.m.add_command(label ='변화율 지정/해제', command=lambda:self.toChangeRate())
        self.m.add_command(label ='순서 변경', command=lambda:self.changePosition())





################################################################################################################################################################################################################




        # 날짜 입력 프레임
        self.dateFrame = tk.Frame(self)
        self.dateFrame.grid(row=0, column=1, padx=5)

        #시작일
        yearList = [str(i)+'년' for i in range(1980, today.year+1)]     
        self.startYearCombo = tkCombo.Combobox(self.dateFrame, height=0, values=yearList, state='readonly')
        self.startYearCombo.current(len(yearList)-1)
        self.startYearCombo.grid(row=0, column=0, padx= 5)

        monthList = [str(i)+'월' for i in range(1, 13)]
        self.startMonthCombo = tkCombo.Combobox(self.dateFrame, height=0, values=monthList, state='readonly')
        self.startMonthCombo.current(today.month-1)
        self.startMonthCombo.bind("<<ComboboxSelected>>", lambda event : self.selectDate(event, 'startMonth'))
        self.startMonthCombo.grid(row=0, column=1, padx= 5)

        self.startDayValue = [str(i)+'일' for i in range(1, lastDay+1)]
        self.startDayCombo = tkCombo.Combobox(self.dateFrame, height=0, values=self.startDayValue, state='readonly')
        self.startDayCombo.current(0)
        self.startDayCombo.grid(row=0, column=2, padx= 5)
        
        tk.Label(self.dateFrame, text='에서').grid(row=0, column=3, padx= 5, sticky=tk.W)

        #종료일
        self.endYearCombo = tkCombo.Combobox(self.dateFrame, height=0, values=yearList, state='readonly')
        self.endYearCombo.current(len(yearList)-1)
        self.endYearCombo.grid(row=1, column=0, padx= 5)

        self.endMonthCombo = tkCombo.Combobox(self.dateFrame, height=0, values=monthList, state='readonly')
        self.endMonthCombo.current(today.month-1)
        self.endMonthCombo.bind("<<ComboboxSelected>>", lambda event:self.selectDate(event, 'endMonth'))
        self.endMonthCombo.grid(row=1, column=1, padx= 5)

        self.endDayValue = [str(i)+'일' for i in range(1, lastDay+1)]
        self.endDayCombo = tkCombo.Combobox(self.dateFrame, height=0, values=self.endDayValue, state='readonly')
        self.endDayCombo.current(today.day-1)
        self.endDayCombo.grid(row=1, column=2, padx= 5)

        tk.Label(self.dateFrame, text='까지').grid(row=1, column=3, padx= 5, sticky=tk.W)



################################################################################################################################################################################################################



        # 리스트 프레임
        self.listFrame = tk.Frame(self)
        self.listFrame.grid(row=1, column=1, padx=5)

        # ecos 리스트박스
        tk.Label(self.listFrame, text='한국은행에서').grid(row=0, column=0)

        self.ecosSelect = tk.Listbox(self.listFrame, selectmode='single', height=20)
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
        ecosScroll = Scrollbar(self.listFrame, orient='vertical')
        self.ecosSelect.config(yscrollcommand=ecosScroll.set)
        ecosScroll.config(command=self.ecosSelect.yview)
        ecosScroll.grid(row=1, column=1, sticky=tk.N + tk.S + tk.W)

        #ecos 리스트박스 이벤트 연결 및 배치
        self.ecosSelect.bind('<Double-Button-1>', lambda event:self.addToTreeview(event, 'ecos'))
        self.ecosSelect.grid(row=1, column=0, sticky=tk.N + tk.E)





        # yf 리스트 박스
        tk.Label(self.listFrame, text='야후파이낸스에서').grid(row=0, column=2)
        self.yfSelect = tk.Listbox(self.listFrame, selectmode='single', height=20)
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
        yfScroll = Scrollbar(self.listFrame, orient='vertical')
        self.yfSelect.config(yscrollcommand=yfScroll.set)
        yfScroll.config(command=self.yfSelect.yview)
        yfScroll.grid(row=1, column=3, sticky=tk.N + tk.S + tk.W)

        #yf 리스트박스 이벤트 연결 및 배치
        self.yfSelect.grid(row=1, column=2, sticky=tk.N + tk.E)
        self.yfSelect.bind('<Double-Button-1>', lambda event:self.addToTreeview(event, 'yf'))





        # fred 리스트 박스
        tk.Label(self.listFrame, text='fred에서').grid(row=0, column=4)
        self.fredSelect = tk.Listbox(self.listFrame, selectmode='single', height=20)
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
        fredScroll = Scrollbar(self.listFrame, orient='vertical')
        self.fredSelect.config(yscrollcommand=fredScroll.set)
        fredScroll.config(command=self.fredSelect.yview)
        fredScroll.grid(row=1, column=5, sticky=tk.N + tk.S + tk.W)

        #fred 리스트박스 이벤트 연결 및 배치
        self.fredSelect.grid(row=1, column=4, sticky=tk.N)
        self.fredSelect.bind('<Double-Button-1>', lambda event:self.addToTreeview(event, 'fred'))





        # krx 프레임
        self.krxFrame = tk.Frame(self.listFrame)
        self.krxFrame.grid(row=0, column=6, rowspan=2, sticky=tk.N)

        # krx 개별종목 코드입력
        tk.Label(self.krxFrame, text='krx에서\n').grid(row=0, column=0)
        self.e= tk.Entry(self.krxFrame, width=20)
        self.e.grid(row=1, column=0, sticky=tk.N)


        self.krx_check_dic = {'PRICE':tk.IntVar(), 'BPS':tk.IntVar(), 'PER':tk.IntVar(), 'PBR':tk.IntVar(), 'EPS':tk.IntVar(), 'DIV':tk.IntVar(), 'DPS':tk.IntVar()}
        
        priceCheck = Checkbutton(self.krxFrame, text='PRICE', variable=self.krx_check_dic['PRICE'])
        priceCheck.grid(row=2, column=0)
        bpsCheck = Checkbutton(self.krxFrame, text='BPS', variable=self.krx_check_dic['BPS'])
        bpsCheck.grid(row=3, column=0)
        perCheck = Checkbutton(self.krxFrame, text='PER', variable=self.krx_check_dic['PER'])
        perCheck.grid(row=4, column=0)
        perCheck = Checkbutton(self.krxFrame, text='PBR', variable=self.krx_check_dic['PBR'])
        perCheck.grid(row=5, column=0)
        perCheck = Checkbutton(self.krxFrame, text='EPS', variable=self.krx_check_dic['EPS'])
        perCheck.grid(row=6, column=0)
        perCheck = Checkbutton(self.krxFrame, text='DIV', variable=self.krx_check_dic['DIV'])
        perCheck.grid(row=7, column=0)
        perCheck = Checkbutton(self.krxFrame, text='DPS', variable=self.krx_check_dic['DPS'])
        perCheck.grid(row=8, column=0)
        # self.krxBtn = tk.Button(self.krxFrame, text='목록에 추가', command=lambda event:self.addToTreeview(event, 'krx'))     # 이렇게하니 버튼 이벤트 안 됨.
        self.krxBtn = tk.Button(self.krxFrame, text='목록에 추가')
        self.krxBtn.bind('<Button-1>', lambda event:self.addToTreeview(event, 'krx'))
        self.krxBtn.grid(row=9, column=0)





################################################################################################################################################################################################################





        # 입력 프레임
        self.inputFrame = tk.Frame(self)
        self.inputFrame.grid(row=2, column=1)

        # 기준 주기 콤보
        self.spList = [
            {'name':'원자료 주기', 'sp':None}, 
            {'name':'QQ(분기)', 'sp':'QQ'}, 
            {'name':'MM(월)', 'sp':'MM'},
            {'name':'DD(일)', 'sp':'DD'},
        ]
        self.spCombo = tkCombo.Combobox(self.inputFrame, height=0, width=50, values=[item['name'] for item in self.spList], state='readonly')
        self.spCombo.current(0)
        self.spCombo.grid(row=0, column=0)
       
        # 서버에서 가져오기 버튼
        self.btnRun = tk.Button(self.inputFrame, text='가져오기', command=lambda:self.getData())
        self.btnRun.grid(row=0, column=1, padx=5)

        # 불러오기 버튼
        self.btnRun = tk.Button(self.inputFrame, text='불러오기', command=lambda:self.loadData())
        self.btnRun.grid(row=1, column=1, pady=10)
        self.picklePath = tk.Entry(self.inputFrame, width=50)
        self.picklePath.grid(row=1, column=0, padx=5, pady=20)

        # 출력하기 버튼
        self.btnRun = tk.Button(self.inputFrame, text='출력', command=lambda:self.printData())
        self.btnRun.grid(row=2, column=1, padx=5)
        
        # 출력기능 콤보
        self.functionList = [
            {'name':'그래프출력', 'param':'graph'},
            {'name':'엑셀파일로 출력', 'param':'cor'},
            {'name':'수익률 비교', 'param':'ccr'},   # Comparison of Cumulative Returns
            {'name':'히스토그램(누적수익률)', 'param':'histo'},
            {'name':'이동평균분석', 'param':'ma'},
            {'name':'볼린저밴드', 'param':'band'},
        ]
        self.functioinCombo = tkCombo.Combobox(self.inputFrame, height=0, width=50, values=[item['name'] for item in self.functionList], state='readonly')
        self.functioinCombo.current(0)
        self.functioinCombo.grid(row=2, column=0)





################################################################################################################################################################################################################




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

        time.sleep(3)
        browser.find_element(By.ID, 'down').click()
        time.sleep(3)

        browser.quit()





################################################################################################################################################################################################################




    def loadData(self):
        value = filedialog.askopenfile(parent=self, initialdir='./pickles', title='저장된 피클 선택', filetypes=[('pickle', '.*')])
        if value == None:
            return
        else:
            self.picklePath.delete(0, tk.END)
            self.picklePath.insert(0, value.name)





################################################################################################################################################################################################################




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





################################################################################################################################################################################################################




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




################################################################################################################################################################################################################





    def addToTreeview(self, event, widget):
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
            if not(code):
                print('코드를 입력하세요')
                return False
            mode = []
            for key, val in self.krx_check_dic.items():
                if val.get():
                    mode.append(key)    
            dict = {'source':'krx', 'name':code, 'code':code, 'changeRate':False, 'mode':mode}
        macroFrame.selected.append(dict)

        self.updateTree()




################################################################################################################################################################################################################





    def delFromSelected(self):
        items = self.selectedTree.selection()
        itemList = self.selectedTree.get_children()
        
        for item in reversed(items):
            index = itemList.index(item)
            del macroFrame.selected[index]
        self.updateTree()




################################################################################################################################################################################################################





    def toChangeRate(self):
        item = self.selectedTree.focus()
        itemList = self.selectedTree.get_children()
        index = itemList.index(item)
        macroFrame.selected[index]['changeRate'] = not(macroFrame.selected[index]['changeRate'])
        self.updateTree()




################################################################################################################################################################################################################





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




################################################################################################################################################################################################################





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





################################################################################################################################################################################################################





    def rightClickEvent(self, event):
        self.m.tk_popup(event.x_root, event.y_root)
        self.m.grab_release()





################################################################################################################################################################################################################





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
    # macroFrame().pack()
    macroFrame().grid()
    root.mainloop()