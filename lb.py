import sqlite3
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showerror
from tkinter.messagebox import showinfo
import re
import pandas as pd

# описываем класс Окно
class Main_Window():
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("700x550")
        self.root.title("Магазин музыкальных инструментов")
        # логотип размещаем по центру окна
        self.logo = tk.PhotoImage(file="logo.png")
        self.label = tk.Label(self.root, image=self.logo)
        self.label.pack(anchor='center', expand=1)
        # создаем меню
        self.main_menu = tk.Menu()
        self.file_menu = tk.Menu(tearoff=0)  # подменю
        self.ref_menu = tk.Menu(tearoff=0)  # подменю
        self.book_menu = tk.Menu(tearoff=0)  # подменю
        self.otch_menu = tk.Menu(tearoff=0)  # подменю
        self.help_menu = tk.Menu(tearoff=0)  # подменю
        self.file_menu.add_command(label="Выход", command=quit)
        self.ref_menu.add_command(label="Поставщики", command=self.open_win_provider)
        self.ref_menu.add_command(label="Музыкальные Инструменты", command=self.open_win_name)
        self.ref_menu.add_command(label="Виды инструментов",command=self.open_win_types)
        self.book_menu.add_command(label='Поступление музыкальных инструментов')
        self.book_menu.add_command(label='Список музыкальных инструментов',command=self.open_win_book)
        self.book_menu.add_command(label='Продажа музыкальных инструментов')
        self.otch_menu.add_command(label="Отчет по поступлению")
        self.otch_menu.add_command(label="Отчет по остаткам")
        self.otch_menu.add_command(label="Отчет по продажам")
        self.help_menu.add_command(label="Руководство пользователя")
        self.help_menu.add_command(label="О программе")
        self.main_menu.add_cascade(label="Файл", menu=self.file_menu)
        self.main_menu.add_cascade(label="Справочники", menu=self.ref_menu)
        self.main_menu.add_cascade(label="Музыкальные инструменты", menu=self.book_menu)
        self.main_menu.add_cascade(label="Отчеты", menu=self.otch_menu)
        self.main_menu.add_cascade(label="Справка", menu=self.help_menu)
        # привязываем меню к созданному окну
        self.root.config(menu=self.main_menu)
    def open_win_provider(self):
        self.root.withdraw()  # скрыть окно
        Provider_Window()
    def open_win_name(self):
        self.root.withdraw()  # скрыть окно
        Name_Window()
    def open_win_types(self):
        self.root.withdraw()  # скрыть окно
        Types_Window()
    def open_win_book(self):
        '''метод Открыть окно Книги'''
        self.root.withdraw()  # скрыть окно
        Book_window()

class Provider_Window():
    '''Окно Поставщики'''
    def __init__(self):
        self.root2 = tk.Tk()
        self.root2.geometry("800x500")
        self.root2.title("Магазин музыкальных инструментов/Поставщики")
        self.root2.protocol('WM_DELETE_WINDOW', lambda: self.quit_win_provider())  # перехват кнопки Х
        self.main_view = win
        self.db = db
        # фреймы
        self.table_frame = tk.Frame(self.root2, bg='green')
        self.add_edit_frame = tk.Frame(self.root2, bg='red')
        self.table_frame.place(relx=0, rely=0, relheight=1, relwidth=0.6)
        self.add_edit_frame.place(relx=0.6, rely=0, relheight=1, relwidth=0.4)
        # таблица
        self.table_pr = ttk.Treeview(self.table_frame, columns=('name_provider', 'address', 'contacts'), height=15, show='headings')
        self.table_pr.column("name_provider", width=150, anchor=tk.NW)
        self.table_pr.column("address", width=200, anchor=tk.NW)
        self.table_pr.column("contacts", width=120, anchor=tk.CENTER)
        self.table_pr.heading("name_provider", text='ФИО')
        self.table_pr.heading("address", text='Адрес')
        self.table_pr.heading("contacts", text='Контактная информация')
        self.view_info()
        # Полоса прокрутки
        self.scroll_bar = ttk.Scrollbar(self.table_frame)
        self.table_pr['yscrollcommand'] = self.scroll_bar.set
        self.scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)
        self.table_pr.place(relx=0, rely=0, relheight=0.9, relwidth=0.97)
        # поле ввода и кнопка для поиска
        self.esearch = ttk.Entry(self.table_frame)
        self.esearch.place(relx=0.02, rely=0.92, relheight=0.05, relwidth=0.7)
        self.butsearch = tk.Button(self.table_frame, text="Найти по ФИО", command=self.search_info)
        self.butsearch.place(relx=0.74, rely=0.92, relheight=0.05, relwidth=0.2)
        # поля для ввода
        self.lname = tk.Label(self.add_edit_frame, text="ФИО")
        self.lname.place(relx=0.04, rely=0.02, relheight=0.05, relwidth=0.4)
        self.ename = ttk.Entry(self.add_edit_frame)
        self.ename.place(relx=0.45, rely=0.02, relheight=0.05, relwidth=0.5)
        self.lcontact = tk.Label(self.add_edit_frame, text="Адрес")
        self.lcontact.place(relx=0.04, rely=0.12, relheight=0.05, relwidth=0.4)
        self.econtact = ttk.Entry(self.add_edit_frame)
        self.econtact.place(relx=0.45, rely=0.12, relheight=0.05, relwidth=0.5)
        self.lphone = tk.Label(self.add_edit_frame, text="Конт-ная информация")
        self.lphone.place(relx=0.04, rely=0.22, relheight=0.05, relwidth=0.4)
        # валидация номера телефона для поля ввода
        self.ephone = ttk.Entry(self.add_edit_frame)
        self.ephone.place(relx=0.45, rely=0.22, relheight=0.05, relwidth=0.5)
        self.ephone.insert(0, "+375")
        # кнопки
        self.butadd = tk.Button(self.add_edit_frame, text="Добавить запись",command=self.add_info)
        self.butadd.place(relx=0.1, rely=0.33, relheight=0.07, relwidth=0.8)
        self.butdel = tk.Button(self.add_edit_frame, text="Удалить запись", command=self.delete_info)
        self.butdel.place(relx=0.1, rely=0.44, relheight=0.07, relwidth=0.8)
        self.buted = tk.Button(self.add_edit_frame, text="Редактировать запись", command=self.update_info)
        self.buted.place(relx=0.1, rely=0.55, relheight=0.07, relwidth=0.8)
        self.butsave = tk.Button(self.add_edit_frame, text="Сохранить изменения", command=self.save_info)
        self.butsave.place(relx=0.1, rely=0.66, relheight=0.07, relwidth=0.8)
        self.butquit = tk.Button(self.add_edit_frame, text="Закрыть", command=self.quit_win_provider)
        self.butquit.place(relx=0.1, rely=0.77, relheight=0.07, relwidth=0.8)
    def view_info(self):
        '''отобразить данные таблицы Поставщики'''
        self.db.c.execute('''SELECT name_provider, adress, contact_information FROM provider''')
        [self.table_pr.delete(i) for i in self.table_pr.get_children()]  # очистить таблицу для последующего обновления
        [self.table_pr.insert('', 'end', values=row) for row in self.db.c.fetchall()]
    def add_info(self):
        '''Добавить запись'''
        info_list_pr = [self.ename.get(), self.econtact.get(), self.ephone.get()]
        result = re.match("^\+{0,1}\d{0,12}$", info_list_pr[2])
        if '' in info_list_pr:  # валидация заполненя всех полей ввода
            showerror(title="Ошибка", message="Все поля должны быть заполнены")
        elif not result or len(info_list_pr[2]) != 13:  # валидация номера телефона для поля ввода
            showerror(title="Ошибка", message="Номер телефона должен быть в формате +375xxxxxxxxx, где x представляет цифру")
        else:
            self.db.c.execute('''INSERT INTO provider(name_provider, adress, contact_information) VALUES (?, ?, ?)''', (info_list_pr[0], info_list_pr[1], info_list_pr[2]))
            self.db.conn.commit()
            self.view_info()
            self.ename.delete(0, tk.END)
            self.econtact.delete(0, tk.END)
            self.ephone.delete(4, tk.END)
    def delete_info(self):
        '''Удалить запись'''
        try:
            self.db.c.execute('''DELETE FROM provider WHERE name_provider=? AND adress=? AND contact_information=?''',
                              (self.table_pr.item(self.table_pr.selection())['values'][0],
                               self.table_pr.item(self.table_pr.selection())['values'][1],
                               '+' + str(self.table_pr.item(self.table_pr.selection())['values'][2])))
            self.db.conn.commit()
            self.view_info()
        except IndexError:
            showinfo(title="Внимание!", message="Выберите запись для удаления")
    def update_info(self):
        '''Редактировать запись'''
        try:
            self.ename.delete(0, tk.END)
            self.ename.insert(0, self.table_pr.item(self.table_pr.selection())['values'][0])
            self.econtact.delete(0, tk.END)
            self.econtact.insert(0, self.table_pr.item(self.table_pr.selection())['values'][1])
            self.ephone.delete(0, tk.END)
            self.ephone.insert(0, '+' + str(self.table_pr.item(self.table_pr.selection())['values'][2]))
        except IndexError:
            showinfo(title="Внимание!", message="Выберите запись для редактирования")
    def save_info(self):
        '''Сохранить изменения'''
        self.db.c.execute('''SELECT * FROM provider''')
        provider_info = self.db.c.fetchall()
        for el in provider_info:
            if el[1] == self.table_pr.item(self.table_pr.selection())['values'][0] and el[2] == \
                    self.table_pr.item(self.table_pr.selection())['values'][1] and el[3] == '+' + str(
                    self.table_pr.item(self.table_pr.selection())['values'][2]):
                provider_id = el[0]
                break
        info_list_pr = [self.ename.get(), self.econtact.get(), self.ephone.get()]
        result = re.match("^\+{0,1}\d{0,12}$", info_list_pr[2])
        if '' in info_list_pr:  # валидация заполненя всех полей ввода
            showerror(title="Ошибка", message="Все поля должны быть заполнены")
        elif not result or len(info_list_pr[2]) != 13:  # валидация номера телефона для поля ввода
            showerror(title="Ошибка", message="Номер телефона должен быть в формате +375xxxxxxxxx, где x представляет цифру")
        else:
            self.db.c.execute('''UPDATE provider SET name_provider=?, adress=?, contact_information=? WHERE id_provider=?''', (info_list_pr[0], info_list_pr[1], info_list_pr[2], provider_id))
        self.db.conn.commit()
        self.view_info()
        self.ename.delete(0, tk.END)
        self.econtact.delete(0, tk.END)
        self.ephone.delete(4, tk.END)
    def search_info(self):
        '''Кнопка Найти'''
        info = ('%' + self.esearch.get() + '%',)
        self.db.c.execute('''SELECT name_provider, adress, contact_information FROM provider 
                          WHERE name_provider LIKE ?''', info)
        [self.table_pr.delete(i) for i in self.table_pr.get_children()]  # очистить таблицу для последующего обновления
        [self.table_pr.insert('', 'end', values=row) for row in self.db.c.fetchall()]
    def quit_win_provider(self):
        self.root2.destroy()
        self.main_view.root.deiconify()

class Name_Window():
    '''Окно Музыкальный Инструмент'''
    def __init__(self):
        self.root2 = tk.Tk()
        self.root2.geometry("800x500")
        self.root2.title("Магазин музыкальных инструментов/Музыкальные Инструменты")
        self.root2.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())  # перехват кнопки Х
        self.main_view = win
        self.db = db
        # фреймы
        self.table_frame = tk.Frame(self.root2, bg='green')
        self.add_edit_frame = tk.Frame(self.root2, bg='red')
        self.table_frame.place(relx=0, rely=0, relheight=1, relwidth=0.6)
        self.add_edit_frame.place(relx=0.6, rely=0, relheight=1, relwidth=0.4)
        # таблица
        self.table_pr = ttk.Treeview(self.table_frame, columns=('name', 'price', 'id_type'), height=15, show='headings')
        self.table_pr.column("name", width=150, anchor=tk.NW)
        self.table_pr.heading("name", text='Название')
        self.table_pr.column("price", width=150, anchor=tk.NW)
        self.table_pr.heading("price", text='Цена')
        self.table_pr.column("id_type", width=150, anchor=tk.NW)
        self.table_pr.heading("id_type", text='Id')
        self.view_info()
        # Полоса прокрутки
        self.scroll_bar = ttk.Scrollbar(self.table_frame, command=self.table_pr.yview)
        self.table_pr['yscrollcommand'] = self.scroll_bar.set
        self.scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)
        self.table_pr.place(relx=0, rely=0, relheight=0.9, relwidth=0.97)
        # поле ввода и кнопка для поиска
        self.esearch = ttk.Entry(self.table_frame)
        self.esearch.place(relx=0.02, rely=0.92, relheight=0.05, relwidth=0.7)
        self.butsearch = tk.Button(self.table_frame, text="Найти", command=self.search_info)
        self.butsearch.place(relx=0.74, rely=0.92, relheight=0.05, relwidth=0.2)
        # поля для ввода
        self.lname = tk.Label(self.add_edit_frame, text="Название")
        self.lname.place(relx=0.04, rely=0.02, relheight=0.05, relwidth=0.4)
        self.ename = ttk.Entry(self.add_edit_frame)
        self.ename.place(relx=0.45, rely=0.02, relheight=0.05, relwidth=0.5)
        self.lprice = tk.Label(self.add_edit_frame, text="Цена")
        self.lprice.place(relx=0.04, rely=0.12, relheight=0.05, relwidth=0.4)
        self.eprice = ttk.Entry(self.add_edit_frame)
        self.eprice.place(relx=0.45, rely=0.12, relheight=0.05, relwidth=0.5)
        self.ltype = tk.Label(self.add_edit_frame, text="Id")
        self.ltype.place(relx=0.04, rely=0.22, relheight=0.05, relwidth=0.4)
        self.etype = ttk.Entry(self.add_edit_frame)
        self.etype.place(relx=0.45, rely=0.22, relheight=0.05, relwidth=0.5)
        # кнопки
        self.butadd = tk.Button(self.add_edit_frame, text="Добавить запись", command=self.add_info)
        self.butadd.place(relx=0.1, rely=0.33, relheight=0.07, relwidth=0.8)
        self.butdel = tk.Button(self.add_edit_frame, text="Удалить запись", command=self.delete_info)
        self.butdel.place(relx=0.1, rely=0.44, relheight=0.07, relwidth=0.8)
        self.buted = tk.Button(self.add_edit_frame, text="Редактировать запись", command=self.update_info)
        self.buted.place(relx=0.1, rely=0.55, relheight=0.07, relwidth=0.8)
        self.butsave = tk.Button(self.add_edit_frame, text="Сохранить изменения", command=self.save_info)
        self.butsave.place(relx=0.1, rely=0.66, relheight=0.07, relwidth=0.8)
        self.butquit = tk.Button(self.add_edit_frame, text="Закрыть", command=self.quit_win)
        self.butquit.place(relx=0.1, rely=0.77, relheight=0.07, relwidth=0.8)
    def view_info(self):
        '''отобразить данные таблицы Места издания'''
        self.db.c.execute('''SELECT name, price, id_type FROM musical_instrument''')
        [self.table_pr.delete(i) for i in self.table_pr.get_children()]  # очистить таблицу для последующего обновления
        [self.table_pr.insert('', 'end', values=row) for row in self.db.c.fetchall()]
    def add_info(self):
        '''Добавить запись'''
        info_list_pr = [self.ename.get(), self.eprice.get(), self.etype.get()]
        self.db.c.execute('''INSERT INTO musical_instrument(name, price, id_type) VALUES (?, ?, ?)''', (info_list_pr[0], info_list_pr[1], info_list_pr[2]))
        self.db.conn.commit()
        self.view_info()
        self.ename.delete(0, tk.END)
        self.eprice.delete(0, tk.END)
        self.etype.delete(0, tk.END)
    def delete_info(self):
        '''Удалить запись'''
        try:
            self.db.c.execute('''DELETE FROM musical_instrument WHERE name=? AND price=? AND id_type=?''',
                          (self.table_pr.item(self.table_pr.selection())['values'][0],
                           self.table_pr.item(self.table_pr.selection())['values'][1],
                           '+' + str(self.table_pr.item(self.table_pr.selection())['values'][2])))
            self.db.conn.commit()
            self.view_info()
        except IndexError:
            showinfo(title="Внимание!", message="Выберите запись для удаления")
    def update_info(self):
        '''Редактировать запись'''
        try:
            self.ename.delete(0, tk.END)
            self.ename.insert(0, self.table_pr.item(self.table_pr.selection())['values'][0])
            self.eprice.delete(0, tk.END)
            self.eprice.insert(0, self.table_pr.item(self.table_pr.selection())['values'][1])
            self.etype.delete(0, tk.END)
            self.etype.insert(0, self.table_pr.item(self.table_pr.selection())['values'][2])
        except IndexError:
            showinfo(title="Внимание!", message="Выберите запись для редактирования")
    def save_info(self):
        '''Сохранить изменения'''
        self.db.c.execute('''SELECT * FROM musical_instrument''')
        instrument_info = self.db.c.fetchall()
        for el in instrument_info:
            if el[1] == self.table_pr.item(self.table_pr.selection())['values'][0] and el[2] == \
                    self.table_pr.item(self.table_pr.selection())['values'][1] and el[3] == '+' + str(
                self.table_pr.item(self.table_pr.selection())['values'][2]):
                instrument_id = el[0]
                break
        info_list_pr = [self.ename.get(), self.eprice.get(), self.etype.get()]
        if '' in info_list_pr:  # валидация заполненя всех полей ввода
            showerror(title="Ошибка", message="Все поля должны быть заполнены")
        else:
            self.db.c.execute(
                '''UPDATE musical_instrument SET name=?, price=? WHERE id_type=?''',
                (info_list_pr[0], info_list_pr[1], info_list_pr[2]))
        self.db.conn.commit()
        self.view_info()
        self.ename.delete(0, tk.END)
        self.eprice.delete(0, tk.END)
        self.etype.delete(4, tk.END)
    def search_info(self):
        '''Кнопка Найти'''
        info = ('%' + self.esearch.get() + '%',)
        self.db.c.execute('''SELECT name, price, id_type FROM musical_instrument 
                          WHERE name LIKE ?''', info)
        [self.table_pr.delete(i) for i in self.table_pr.get_children()]  # очистить таблицу для последующего обновления
        [self.table_pr.insert('', 'end', values=row) for row in self.db.c.fetchall()]
    def quit_win(self):
        self.root2.destroy()
        self.main_view.root.deiconify()

class Types_Window(Provider_Window):
    '''Окно Вид Инструмента'''
    def __init__(self):
        self.root2 = tk.Tk()
        self.root2.geometry("800x500")
        self.root2.title("Магазин музыкальных инструментов/Виды Инструментов")
        self.root2.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())  # перехват кнопки Х
        self.main_view = win
        self.db = db
        # фреймы
        self.table_frame = tk.Frame(self.root2, bg='green')
        self.add_edit_frame = tk.Frame(self.root2, bg='red')
        self.table_frame.place(relx=0, rely=0, relheight=1, relwidth=0.6)
        self.add_edit_frame.place(relx=0.6, rely=0, relheight=1, relwidth=0.4)
        # таблица
        self.table_pr = ttk.Treeview(self.table_frame, columns=('type'), height=15, show='headings')
        self.table_pr.column("type", width=150, anchor=tk.NW)
        self.table_pr.heading("type", text='Виды Инструментов')
        self.view_info()
        # Полоса прокрутки
        self.scroll_bar = ttk.Scrollbar(self.table_frame, command=self.table_pr.yview)
        self.table_pr['yscrollcommand'] = self.scroll_bar.set
        self.scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)
        self.table_pr.place(relx=0, rely=0, relheight=0.9, relwidth=0.97)
        # поле ввода и кнопка для поиска
        self.esearch = ttk.Entry(self.table_frame)
        self.esearch.place(relx=0.02, rely=0.92, relheight=0.05, relwidth=0.7)
        self.butsearch = tk.Button(self.table_frame, text="Найти", command=self.search_info)
        self.butsearch.place(relx=0.74, rely=0.92, relheight=0.05, relwidth=0.2)
        # поля для ввода
        self.ltype = tk.Label(self.add_edit_frame, text="Вид Инструмента")
        self.ltype.place(relx=0.04, rely=0.02, relheight=0.05, relwidth=0.4)
        self.etype = ttk.Entry(self.add_edit_frame)
        self.etype.place(relx=0.45, rely=0.02, relheight=0.05, relwidth=0.5)
        # кнопки
        self.butadd = tk.Button(self.add_edit_frame, text="Добавить запись", command=self.add_info)
        self.butadd.place(relx=0.1, rely=0.33, relheight=0.07, relwidth=0.8)
        self.butdel = tk.Button(self.add_edit_frame, text="Удалить запись", command=self.delete_info)
        self.butdel.place(relx=0.1, rely=0.44, relheight=0.07, relwidth=0.8)
        self.buted = tk.Button(self.add_edit_frame, text="Редактировать запись", command=self.update_info)
        self.buted.place(relx=0.1, rely=0.55, relheight=0.07, relwidth=0.8)
        self.butsave = tk.Button(self.add_edit_frame, text="Сохранить изменения", command=self.save_info)
        self.butsave.place(relx=0.1, rely=0.66, relheight=0.07, relwidth=0.8)
        self.butquit = tk.Button(self.add_edit_frame, text="Закрыть", command=self.quit_win)
        self.butquit.place(relx=0.1, rely=0.77, relheight=0.07, relwidth=0.8)
    def view_info(self):
        '''отобразить данные таблицы Вид инструмента'''
        self.db.c.execute('''SELECT type FROM type''')
        [self.table_pr.delete(i) for i in self.table_pr.get_children()]  # очистить таблицу для последующего обновления
        [self.table_pr.insert('', 'end', values=row) for row in self.db.c.fetchall()]
    def add_info(self):
        '''Добавить запись'''
        self.db.c.execute('''INSERT INTO type(type) VALUES (?)''', (self.etype.get(),))
        self.db.conn.commit()
        self.view_info()
        self.etype.delete(0, tk.END)
    def delete_info(self):
        '''Удалить запись'''
        try:
            selected_value = self.table_pr.item(self.table_pr.selection())['values'][0]
            self.db.c.execute('''DELETE FROM type WHERE type=?''', (selected_value,))
            self.db.conn.commit()
            self.view_info()
        except IndexError:
            showinfo(title="Внимание!", message="Выберите запись для удаления")
    def update_info(self):
        '''Редактировать запись'''
        try:
            self.etype.delete(0, tk.END)
            self.etype.insert(0, self.table_pr.item(self.table_pr.selection())['values'][0])
        except IndexError:
            showinfo(title="Внимание!", message="Выберите запись для редактирования")

    def save_info(self):
        '''Сохранить изменения'''
        self.db.c.execute('''SELECT * FROM type''')
        instrument_info = self.db.c.fetchall()
        for el in instrument_info:
            if (
                    el[1] == self.table_pr.item(self.table_pr.selection())['values'][0]
            ):
                instrument_id = el[0]
                break
        info_list_pr = [self.etype.get()]
        if '' in info_list_pr:  # валидация заполненя всех полей ввода
            showerror(title="Ошибка", message="Все поля должны быть заполнены")
        else:
            self.db.c.execute('''UPDATE type SET type=? WHERE id_type=?''', (info_list_pr[0], instrument_id))
        self.db.conn.commit()
        self.view_info()
        self.etype.delete(4, tk.END)
    def search_info(self):
        '''Кнопка Найти'''
        info = ('%' + self.esearch.get() + '%',)
        self.db.c.execute('''SELECT type FROM type 
                          WHERE type LIKE ?''', info)
        [self.table_pr.delete(i) for i in self.table_pr.get_children()]  # очистить таблицу для последующего обновления
        [self.table_pr.insert('', 'end', values=row) for row in self.db.c.fetchall()]
    def quit_win(self):
        self.root2.destroy()
        self.main_view.root.deiconify()


class Book_window():
    def __init__(self):
        self.root2 = tk.Tk()
        self.root2.geometry("800x400")
        self.root2.title("Магазин музыкальных инструментов/Музыкальные Инструменты")
        self.root2.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())  # перехват кнопки Х
        self.main_view = win
        self.db = db
        self.label = tk.Label(self.root2, text="Музыкальные Инструменты")
        self.label.pack(anchor='nw')
        self.button_toexcel = tk.Button(self.root2, text="Экспорт в Excel", command=self.toexcel_book)
        self.button_toexcel.pack(anchor='nw', expand=1)
        self.tree = ttk.Treeview(self.root2, columns=('id_instrument', 'name', 'type', 'price'), height=15, show='headings')
        self.tree.column("id_instrument", width=100, anchor=tk.NW)
        self.tree.column("name", width=100, anchor=tk.NW)
        self.tree.column("type", width=110, anchor=tk.CENTER)
        self.tree.column("price", width=100, anchor=tk.CENTER)
        self.tree.heading("id_instrument", text='Id_инструмента')
        self.tree.heading("name", text='Название')
        self.tree.heading("type", text='Id_вид_инструмента')
        self.tree.heading("price", text='Цена (руб.)')
        # Полоса прокрутки
        self.scroll_bar = ttk.Scrollbar(self.root2, command=self.tree.yview)
        self.tree['yscrollcommand'] = self.scroll_bar.set
        self.scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack()
        self.button_quit = tk.Button(self.root2, text="Закрыть", command=lambda: self.quit_win())
        self.button_quit.pack(anchor='sw', expand=1)
        self.view_records_book()
    def record(self, id_instrument, name, type, price):
        '''обновление и вызов функции для отображения данных'''
        self.db.save_data_book(id_instrument, name, type, price)
        self.view_records_book()
    def view_records_book(self):
        '''отобразить данные таблицы Книги'''
        self.db.c.execute('''SELECT id_instrument, name, id_type, price FROM musical_instrument''')
        [self.tree.delete(i) for i in self.tree.get_children()]  # очистить таблицу для последующего обновления
        [self.tree.insert('', 'end', values=row) for row in self.db.c.fetchall()]
    def update_records_book(self, name, type, price):
        self.db.c.execute('''UPDATE book SET id_instrument=?, name=?, id_type=?, price=?, WHERE ID=?''',
                          (id_instrument, name, type, price, self.tree.set(self.tree.selection()[0], '#1')))
        self.db.conn.commit()
        self.view_records_book()
    def quit_win(self):
        self.root2.destroy()
        self.main_view.root.deiconify()
    def toexcel_book(self):
        self.db.c.execute('''SELECT * FROM musical_instrument''')
        book_list = self.db.c.fetchall()
        # Используем словарь для заполнения DataFrame
        # Ключи в словаре — это названия колонок. А значения - строки с информацией
        df = pd.DataFrame({'Id_инструмента': [el[1] for el in book_list],
                           'Название': [el[2] for el in book_list],
                           'Цена (руб.)': [el[3] for el in book_list],
                           'Вид инструмента': [el[8] for el in book_list]})
        # указажем writer библиотеки
        writer = pd.ExcelWriter('example.xlsx', engine='xlsxwriter')
        # записшем наш DataFrame в файл
        df.to_excel(writer, 'Sheet1')
        # сохраним результат
        writer.save()


class DB:
    def __init__(self):
        self.conn = sqlite3.connect('book_bd.db')  # установили связь с БД (или создали если ее нет)
        self.c = self.conn.cursor()  # создали курсор
        # таблица Книги
        self.c.execute(
            '''CREATE TABLE IF NOT EXISTS "musical_instrument" (
                       "id_instrument" INTEGER NOT NULL,
                        "name" TEXT NOT NULL,
                        "price" REAL NOT NULL,
                        "id_type" INTEGER NOT NULL,
                        PRIMARY KEY("id_instrument" AUTOINCREMENT)
                        )'''
        )
        self.c.execute(
            '''CREATE TABLE IF NOT EXISTS "type" (
                       "id_type" INTEGER NOT NULL,
                        "type" TEXT NOT NULL,
                        PRIMARY KEY("id_type" AUTOINCREMENT)
                        )'''
        )
        # таблица поставщик
        self.c.execute(
            '''CREATE TABLE IF NOT EXISTS "provider" (
                        "id_provider" INTEGER NOT NULL,
                        "name_provider" TEXT NOT NULL,
                        "adress" TEXT NOT NULL,
                        "contact_information" TEXT NOT NULL,
                        PRIMARY KEY("id_provider" AUTOINCREMENT)
                        )'''
        )
        # таблица продажа
        self.c.execute(
            '''CREATE TABLE IF NOT EXISTS "sale" (
                        "id_sale" INTEGER NOT NULL,
                        "date" TEXT NOT NULL,
                        "count" INTEGER NOT NULL,
                        "sum" INTEGER NOT NULL,
                        "id_instrument" INTEGER NOT NULL,
                        PRIMARY KEY("id_sale" AUTOINCREMENT)
                        )'''
        )
        # таблица поступление
        self.c.execute(
            '''CREATE TABLE IF NOT EXISTS "postuplenie" (
                        "id_postuplenie" INTEGER NOT NULL,
                        "date" TEXT NOT NULL,
                        "count" INTEGER NOT NULL,
                        "sum" INTEGER NOT NULL,
                        "id_instrument" INTEGER NOT NULL,
                        "id_provider" INTEGER NOT NULL,
                        PRIMARY KEY("id_postuplenie" AUTOINCREMENT)
                        )'''
        )
        self.conn.commit()
db = DB()
# создаем окно
win = Main_Window()
# запускаем окно
win.root.mainloop()
