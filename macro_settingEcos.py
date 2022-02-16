import tkinter as tk
from tkinter import ttk
import tkinter.ttk as myCombo
import sqlite3

con = sqlite3.connect('./db/myDb.db')
with con:
    cursor = con.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS ecos_indicator(id INTEGER PRIMARY KEY AUTOINCREMENT not null, name TEXT, code1 TEXT, period TEXT, code2_1 TEXT, code2_2 TEXT, code2_3 TEXT)")
    con.commit()





class ecosSetting(tk.Frame):

    def __init__(self):
        
        super().__init__()
        
        # 저장된 계정과목을 출력하는 트리뷰
        self.indicatorTree = ttk.Treeview(self, column=('c0', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6'), show='headings', height=20)
        self.indicatorTree.heading('c0', text='id', anchor=tk.CENTER)
        self.indicatorTree.column('c0', width=50, anchor=tk.CENTER)
        self.indicatorTree.heading('c1', text='name', anchor=tk.CENTER)
        self.indicatorTree.column('c1', width=200, anchor=tk.CENTER)
        self.indicatorTree.heading('c2', text='code1', anchor=tk.CENTER)
        self.indicatorTree.column('c2', width=100, anchor=tk.CENTER)
        self.indicatorTree.heading('c3', text='period', anchor=tk.CENTER)
        self.indicatorTree.column('c3', width=50, anchor=tk.CENTER)
        self.indicatorTree.heading('c4', text='code2_1', anchor=tk.CENTER)
        self.indicatorTree.column('c4', width=100, anchor=tk.CENTER)
        self.indicatorTree.heading('c5', text='code2_2', anchor=tk.CENTER)
        self.indicatorTree.column('c5', width=100, anchor=tk.CENTER)
        self.indicatorTree.heading('c6', text='code2_3', anchor=tk.CENTER)
        self.indicatorTree.column('c6', width=100, anchor=tk.CENTER)
        
        #우측 버튼 클릭시 메뉴 이벤트 연결
        self.indicatorTree.bind('<Button-3>', lambda event:self.rightClick(event))

        #우클릭 이벤트 메뉴
        self.m = tk.Menu(self, tearoff = 0)
        self.m.add_command(label ='추가', command=lambda :self.menuEvent('add'))
        self.m.add_command(label ='수정', command=lambda :self.menuEvent('edit'))
        self.m.add_command(label ='삭제', command=lambda :self.menuEvent('del'))

        #id 컬럼 숨김
        self.indicatorTree['displaycolumns']=('c1','c2','c3','c4','c5','c6')

        #트리뷰 배치
        self.indicatorTree.grid(row=0, column=0, padx=1, pady=5, columnspan=3)

        #스크롤바
        self.yscrollbar = ttk.Scrollbar(self, orient='vertical', command=self.indicatorTree.yview)
        self.yscrollbar.grid(row=0, column=3, sticky='nse')
        self.indicatorTree.configure(yscrollcommand=self.yscrollbar.set)
        self.yscrollbar.configure(command=self.indicatorTree.yview)


        # db에서 저장된 계정과목 가져와서, 트리뷰에 삽입하기
        with con:
            cursor = con.cursor()
            sql = 'select * from ecos_indicator order by name'
            cursor.execute(sql)
            rows = cursor.fetchall()
            for row in rows:
                self.indicatorTree.insert('', 'end', values=(row[0], row[1], row[2], row[3], '.'+row[4], '.'+row[5], '.'+row[6]))


        tk.Label(self, text='지표이름(필수)').grid(row=1, column=0, padx=1, pady=5)
        self.name = tk.Entry(self, width=30)
        self.name.grid(row=1, column=1, padx=1, pady=5)
        
        tk.Label(self, text='통계표코드(필수)').grid(row=2, column=0, padx=1, pady=5)
        self.code1 = tk.Entry(self, width=30)
        self.code1.grid(row=2, column=1, padx=1, pady=5)

        tk.Label(self, text='주기(필수)').grid(row=3, column=0, padx=1, pady=5)
        self.period = myCombo.Combobox(self, height=5, state='readonly', values=['DD', 'MM', 'QQ'])
        self.period.grid(row=3, column=1, padx=1, pady=5)
        
        tk.Label(self, text='통계항목1코드(선택)').grid(row=4, column=0, padx=1, pady=5)
        self.code2_1 = tk.Entry(self, width=30)
        self.code2_1.grid(row=4, column=1, padx=1, pady=5)
        
        tk.Label(self, text='통계항목2코드(선택)').grid(row=5, column=0, padx=1, pady=5)
        self.code2_2 = tk.Entry(self, width=30)
        self.code2_2.grid(row=5, column=1, padx=1, pady=5)
        
        tk.Label(self, text='통계항목3코드(선택)').grid(row=6, column=0, padx=1, pady=5)
        self.code2_3 = tk.Entry(self, width=30)
        self.code2_3.grid(row=6, column=1, padx=1, pady=5)

        tk.Button(self, text='수정하기', command=lambda:self.editBtn()).grid(row=7, column=0, pady=5, columnspan=2)





    def rightClick(self, event):
        self.m.tk_popup(event.x_root, event.y_root)
        self.m.grab_release()





    def menuEvent(self, param):

        if param == 'add':
            self.clearInput()
            with con:
                cursor = con.cursor()
                cursor.executemany(
                    'insert into ecos_indicator(name, code1, period, code2_1, code2_2, code2_3) values(?, ?, ?, ?, ?, ?)',
                    [
                        ('!!!!!새로운 지표!!!!!', '', '', '', '', '')
                    ]
                )
                con.commit()
            self.updateTree()

        elif param == 'del':
            self.clearInput()
            items = self.indicatorTree.selection()
            for item in items:
                # db에서 삭제
                selectedValue = self.indicatorTree.item(item)['values']
                with con:
                    cursor = con.cursor()
                    sql = 'delete from ecos_indicator where id = ?'
                    cursor.execute(sql, [ (selectedValue[0]) ] )
            self.updateTree()

        elif param == 'edit':
            self.clearInput()
            item = self.indicatorTree.selection()
            if item:
                selectedValue = self.indicatorTree.item(item)['values']
                self.name.insert(0, selectedValue[1])
                self.code1.insert(0, selectedValue[2])
                self.code2_1.insert(0, selectedValue[4].replace('.',''))
                self.code2_2.insert(0, selectedValue[5].replace('.',''))
                self.code2_3.insert(0, selectedValue[6].replace('.',''))
                if selectedValue[3] in ['DD','MM','QQ']:
                    self.period.current(['DD','MM','QQ'].index(selectedValue[3]))



    def clearInput(self):
        self.name.delete(0, tk.END)
        self.code1.delete(0, tk.END)
        self.code2_1.delete(0, tk.END)
        self.code2_2.delete(0, tk.END)
        self.code2_3.delete(0, tk.END)





    def updateTree(self):
        for item in self.indicatorTree.get_children():
            self.indicatorTree.delete(item)
        with con:
            cursor = con.cursor()
            cursor.execute('select * from ecos_indicator order by name')
            rows = cursor.fetchall()
            if rows:
                for row in rows:
                    self.indicatorTree.insert(
                        '', 'end', text='', values=(row[0], row[1], row[2], row[3], '.'+row[4], '.'+row[5], '.'+row[6])
                    )





    def editBtn(self):
        item = self.indicatorTree.selection()
        if item:
            id = self.indicatorTree.item(item)['values'][0]
            name = self.name.get()
            code1 = self.code1.get()
            period = self.period.get()
            code2_1 = self.code2_1.get()
            code2_2 = self.code2_2.get()
            code2_3 = self.code2_3.get()
            with con:
                cursor = con.cursor()
                sql = 'update ecos_indicator set name=?, code1=?, period=?, code2_1=?, code2_2=?, code2_3=? where id=?'
                cursor.executemany(sql, [(name, code1, period, code2_1, code2_2, code2_3, id)])
                con.commit()
            self.updateTree()
            self.clearInput()







if __name__ == '__main__':
    root = tk.Tk()
    ecosSetting().pack()
    root.mainloop()