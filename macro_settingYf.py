import tkinter as tk
from tkinter import Scrollbar, ttk
import tkinter.ttk as myCombo
import sqlite3
import tkinter.messagebox as msgbox

con = sqlite3.connect('./db/myDb.db')

with con:
    sqlite3.Connection
    cursor = con.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS yf_indicator(id integer primary key autoincrement not null, name text, symbol text)")
    con.commit()





class yfSetting(tk.Frame):

    def __init__(self):

        super().__init__()

        self.indicatorTree = ttk.Treeview(self, column=('c0', 'c1', 'c2'), show='headings', height=20)
        self.indicatorTree.heading('c0', text='id', anchor=tk.CENTER)
        self.indicatorTree.column('c0', width=200, anchor=tk.CENTER)
        self.indicatorTree.heading('c1', text='name', anchor=tk.CENTER)
        self.indicatorTree.column('c1', width=200, anchor=tk.CENTER)
        self.indicatorTree.heading('c2', text='symbol', anchor=tk.CENTER)
        self.indicatorTree.column('c2', width=200, anchor=tk.CENTER)

        #우측 버튼 클릭시 메뉴 이벤트 연결
        self.indicatorTree.bind('<Button-3>', lambda event:self.rightClick(event))

        #우클릭 이벤트 메뉴
        self.m = tk.Menu(self, tearoff = 0)
        self.m.add_command(label ='추가', command=lambda :self.menuEvent('add'))
        self.m.add_command(label ='수정', command=lambda :self.menuEvent('edit'))
        self.m.add_command(label ='삭제', command=lambda :self.menuEvent('del'))


        #아이디 컬럼 숨기기
        self.indicatorTree['displaycolumns'] = ('c1', 'c2')

        #트리뷰 배치
        self.indicatorTree.grid(row=0, column=0, padx=5, pady=5, columnspan=3)

        #스크롤바
        self.yscrollbar = ttk.Scrollbar(self, orient='vertical', command=self.indicatorTree.yview)
        self.yscrollbar.grid(row=0, column=3, sticky='nse')
        self.indicatorTree.configure(yscrollcommand=self.yscrollbar.set)
        self.yscrollbar.configure(command=self.indicatorTree.yview)

        #db에서 저장된 계정과목 가져와서, 트리뷰에 삽입하기
        with con:
            sqlite3.Connection
            cursor = con.cursor()
            sql = 'select * from yf_indicator order by name'
            cursor.execute(sql)
            rows = cursor.fetchall()
            for row in rows:
                self.indicatorTree.insert('', 'end', text='', values=(row[0], row[1], row[2]))




        self.nameLabel = tk.Label(self, text='이름(필수)')
        self.nameLabel.grid(row=1, column=0, padx=5, pady=5)
        self.name = tk.Entry(self, width=30)
        self.name.grid(row=1, column=1, padx=5, pady=5)
        
        self.symbolLabel = tk.Label(self, text='심볼(필수)')
        self.symbolLabel.grid(row=2, column=0, padx=5, pady=5)
        self.symbol = tk.Entry(self, width=30)
        self.symbol.grid(row=2, column=1, padx=5, pady=5)

        self.editBtn = tk.Button(self, text='수정하기', command=lambda:self.editBtnEvent())
        self.editBtn.grid(row=3, column=0, padx=5, pady=5, columnspan=3)






    def menuEvent(self, param):

        if param == 'add':    
            with con:
                cursor = con.cursor()
                sql = 'insert into yf_indicator(name, symbol) values(?, ?)'
                cursor.executemany(sql,[('!!!!!새로운 지표!!!!!', '')])
                con.commit()
            self.updateTree()
            # self.invisibleInput()

        elif param == 'del':
            items = self.indicatorTree.selection()
            for item in items:
                # db에서 삭제
                selectedValue = self.indicatorTree.item(item)['values']
                with con:
                    cursor = con.cursor()
                    sql = 'delete from yf_indicator where id=?'
                    cursor.execute(sql, [(selectedValue[0])])
            self.updateTree()

        elif param == 'edit':
            self.clearInput()
            item = self.indicatorTree.selection()
            if item:
                selectedValue = self.indicatorTree.item(item)['values']
                self.name.insert(0, selectedValue[1])
                self.symbol.insert(0, selectedValue[2])



   

    def clearInput(self):
        self.name.delete(0, tk.END)
        self.symbol.delete(0, tk.END)





    def rightClick(self, event):
        self.m.tk_popup(event.x_root, event.y_root)
        self.m.grab_release()




    def updateTree(self):
        for item in self.indicatorTree.get_children():
            self.indicatorTree.delete(item)
        with con:
            cursor = con.cursor()
            cursor.execute('select * from yf_indicator order by name')
            rows = cursor.fetchall()
            if rows:
                for row in rows:
                    self.indicatorTree.insert(
                        '', 'end', text='', values=(row[0], row[1], row[2])
                    )


    


    def editBtnEvent(self):
        item = self.indicatorTree.selection()
        if item:
            id = self.indicatorTree.item(item)['values'][0]
            name = self.name.get()
            symbol = self.symbol.get()
            with con:
                cursor = con.cursor()
                sql = 'update yf_indicator set name=?, symbol=? where id=?'
                cursor.executemany(sql, [(name, symbol, id)])
                con.commit()
            self.updateTree()
            self.clearInput()





if __name__ == '__main__':
    root = tk.Tk()
    yfSetting().pack()
    root.mainloop()