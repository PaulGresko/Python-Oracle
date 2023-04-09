import cx_Oracle
import datetime as dt
import pandas as pd
import tkinter as tk
from tkinter import ttk
import tkcalendar
from openpyxl import load_workbook
from tkinter import messagebox



class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()
        self.db = db
        self.view_records()

    def init_main(self):
        toolbar = tk.Frame(bg='#d7d8e0')
        #toolbar = tk.Frame(bg='#d7d8e0', db=2)
        toolbar.pack(side=tk.TOP, fill = tk.X)

        self.add_img = tk.PhotoImage('add.gif')
        btn_open_dialog = tk.Button(toolbar, text='Insert', width=10,height=5, command=self.open_dialog,bg='#d7d8e1',bd=0,compound=tk.TOP)
        btn_open_dialog.pack(side=tk.LEFT)

        btn_updateDialog = tk.Button(toolbar, text='Update', width=10,height=5, command=self.open_updateDialog, bg='#d7d8e2', bd=0,compound=tk.TOP)
        btn_updateDialog.pack(side=tk.LEFT)

        btn_delete = tk.Button(toolbar, text='Delete', width=10,height=5,command=self.delete_records,bg='#d7d8e3', bd=0,compound=tk.TOP)
        btn_delete.pack(side=tk.LEFT)

        btn_refresh = tk.Button(toolbar, text='Refresh', width=10,height=5,command=self.view_records,bg='#d7d8e3', bd=0,compound=tk.TOP)
        btn_refresh.pack(side=tk.LEFT)

        btn_save = tk.Button(toolbar, text='Save', width=10,height=5,command=self.save_records,bg='#d7d8e3', bd=0,compound=tk.TOP)
        btn_save.pack(side=tk.LEFT)

        self.tree =ttk.Treeview(self, columns=('ID_R', 'Title', 'ID', 'Type', 'Section_Name', 'SDate'),height=15,show='headings')
        self.tree.column('ID_R', width=100, anchor=tk.CENTER)
        self.tree.column('Title', width=100, anchor=tk.CENTER)
        self.tree.column('ID', width=100, anchor=tk.CENTER)
        self.tree.column('Type', width=100, anchor=tk.CENTER)
        self.tree.column('Section_Name', width=100, anchor=tk.CENTER)
        self.tree.column('SDate', width=100, anchor=tk.CENTER)

        self.tree.heading('ID_R',text='ID_R')
        self.tree.heading('Title', text='Title')
        self.tree.heading('ID', text='ID')
        self.tree.heading('Type', text='Type')
        self.tree.heading('Section_Name', text='Section Name')
        self.tree.heading('SDate',text='SDate')

        self.tree.pack()

    def records(self, ID_R, Tintle, ID, Type, Section_Name, SDate):
        self.db.insert_data(ID_R, Tintle, ID, Type, Section_Name, SDate)
        self.view_records()

    def update_record(self, Title, ID, Type, Section_Name, SDate):
        #self.db.cur.execute('update GPN_6307.papers SET Title= ?, ID = ?, Type = ?, Section_Name = ?, SDate= ? where ID_R = ?',(Title, ID, Type, Section_Name, SDate, self.tree.set(self.tree.selection() [0],'#1')))
        self.db.cur.execute('update GPN_6307.papers SET TITLE= :1, ID = :2, TYPE = :3, SECTION_NAME = :4, SDATE= :5 where ID_R =:6',(Title, ID, Type, Section_Name, SDate, self.tree.set(self.tree.selection() [0],'#1')))
        self.db.conn.commit()
        self.view_records()

    def save_records(self):
        self.fn = 'save_papers.xlsx'
        self.wb = load_workbook(self.fn)
        self.ws = self.wb['Лист1']
        self.db.cur.execute('select ID_R, Title, FName, Type, Section_name, SDate from GPN_6307.papers, GPN_6307.participants where GPN_6307.papers.id = GPN_6307.participants.id')
        #self.ws.append('ID_R', 'Title', 'FName', 'Type', 'Section_name', 'SDate')
        [self.ws.append(row) for row in self.db.cur.fetchall()]
        self.wb.save(self.fn)
        self.wb.close()

    def view_records(self): # Показывает список
        self.db.cur.execute('select * from GPN_6307.Papers')
        [self.tree.delete(i) for i in self.tree.get_children()]
        for row in self.db.cur.fetchall():
            self.tree.insert('','end', values= row)
        
        

    def delete_records(self):
        #for selection_item in self.tree.selection():
        if not self.tree.selection():
            messagebox.askokcancel('','Select line')
        else:
            self.db.cur.execute('''DELETE FROM GPN_6307.papers where ID_R=:S''', {':S':self.tree.set(self.tree.selection()[0], '#1')})
            self.db.conn.commit()
            self.view_records()

    def open_dialog(self):
        Child()

    def open_updateDialog(self):
        if not self.tree.selection():
            messagebox.askokcancel('','Select line')
        else:
            Update()


class Child(tk.Toplevel):
    def __init__(self):
        super().__init__(root)
        self.init_child()
        self.view = app
        self.db= db

    def init_child(self):
        self.title('Insert')
        self.geometry('400x350+400+300')
        self.resizable(False,False)

        self.label_ID_R = tk.Label(self, text='ID_R')
        self.label_ID_R.place(x=50, y=50)
        self.label_Title = tk.Label(self, text='Title')
        self.label_Title.place(x=50, y=80)
        self.label_ID = tk.Label(self, text='ID')
        self.label_ID.place(x=50,y=110)
        self.label_Type = tk.Label(self, text='Type')
        self.label_Type.place(x=50,y=140)
        self.label_Section_Name = tk.Label(self, text='Section Name')
        self.label_Section_Name.place(x=50,y=170)
        self.label_SDate = tk.Label(self, text='SDate')
        self.label_SDate.place(x=50,y=200)

        
        db.cur.execute('Select ID from GPN_6307.participants group by ID')
        val = db.cur.fetchall()
        #messagebox.askokcancel(val)


        self.entry_ID_R = ttk.Entry(self)
        self.entry_ID_R.place(x=200, y=50)
        self.entry_Title = ttk.Entry(self)
        self.entry_Title.place(x=200, y=80)
        self.entry_ID = ttk.Combobox(self, values=val) # Здесь должны быть id из participants
        self.entry_ID.place(x=200, y=110)
        self.entry_Type = ttk.Combobox(self,values=[u'P',u'O'])
        self.entry_Type.place(x=200, y=140)
        self.entry_Section_Name = ttk.Entry(self)
        self.entry_Section_Name.place(x=200, y=170)
        self.entry_SDate = tkcalendar.DateEntry(self)
        self.entry_SDate.place(x=200, y=200)



        btn_cancel = ttk.Button(self, text='Close',command=self.destroy) # mb fail
        btn_cancel.place(x=300, y=300)

        self.btn_ok = ttk.Button(self,text = 'Insert')
        self.btn_ok.place(x=220, y=300)
        self.btn_ok.bind('<Button-1>', lambda event: self.view.records(self.entry_ID_R.get(), self.entry_Title.get(), self.entry_ID.get(), self.entry_Type.get(), self.entry_Section_Name.get(), self.entry_SDate.get()))

        self.grab_set()
        self.focus_set()

class Update(Child):
    def __init__(self):
        super().__init__()
        self.init_edit()
        self.view = app
        self.db= db
        self.default_data()

    def init_edit(self):
        self.title('Update Data')
        btn_edit = ttk.Button(self,text='Update')
        btn_edit.place(x=220,y=300)
        btn_edit.bind('<Button-1>', lambda event: self.view.update_record(self.entry_Title.get(), self.entry_ID.get(),self.entry_Type.get(),self.entry_Section_Name.get(),self.entry_SDate.get()))
        self.btn_ok.destroy()
        self.label_ID_R.destroy()
        self.entry_ID_R.destroy()

    def default_data(self):

        self.db.cur.execute('''select * from GPN_6307.papers where ID_R= :S''',{':S':self.view.tree.set(self.view.tree.selection()[0],'#1')})
        row = self.db.cur.fetchone()
        self.entry_Title.insert(0,row[1])
        self.entry_ID.insert(0,row[2])
        self.entry_Type.insert(0,row[3])
        self.entry_Section_Name.insert(0,row[4])
        self.entry_SDate.set_date(row[5])

class DB:
    def __init__(self):
        self.conn = cx_Oracle.connect('System/77864059589@localhost:1521')
        self.cur = self.conn.cursor()
        
    def insert_data(self, ID_R, Tintle, ID, Type, Section_Name, SDate):
        self.cur.execute('insert into GPN_6307.papers VALUES(:1, :2, :3, :4, :5, :6)',(ID_R, Tintle, ID, Type, Section_Name,SDate))
        self.conn.commit()


if __name__ == "__main__":
    root = tk.Tk()
    db = DB()
    app = Main(root)
    app.pack()
    root.title("Papers")
    root.geometry("600x450+300+200")
    root.resizable(False,False)
    root.mainloop()

