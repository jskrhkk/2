import sqlite3, sys, os
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from PIL import Image
import pandas as pd
from tkinter.messagebox import showerror, showinfo

MUZ_INSTR_HEADERS = ["Номер по порядку", "Наименование", "ID Вида инструмента", "Цена"]
POSTAVCHIK_HEADERS = ["Номер по порядку","ФИО"]
POSTEUPLENIE_HEADERS = ["Номер по порядку","Дата","ID Муз инструмента", "Номер ТТН", "ID Поставщика" , "Цена за 1 товар" , "Количество", "Сумма"]
SALE_HEADERS = ["Номер по порядку", "Дaта", "ID Муз инстр", "Цена", "Количество", "Сумма"]
VID_MUZ_INSTR_HEADERS = ["Номер по порядку", "Вид муз инструмента"]

class WindowMain(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Музыкальные инструменты')
        self.last_headers = None


        # Создание фрейма для отображения таблицы
        self.table_frame = ctk.CTkFrame(self, width=700, height=500)
        self.table_frame.grid(row=0, column=0, padx=5, pady=5)

        # Загрузка фона
        bg = ctk.CTkImage(Image.open("logo.png"), size=(700, 500))
        lbl = ctk.CTkLabel(self.table_frame, image=bg, text='', font=("Calibri", 40))
        lbl.place(relwidth=1, relheight=1)

        # Создание меню
        self.menu_bar = tk.Menu(self, background='#555', foreground='white')

        # Меню "Файл"
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Выход", command=self.quit)
        self.menu_bar.add_cascade(label="Файл", menu=file_menu)


 # Меню "Справочники"
        references_menu = tk.Menu(self.menu_bar, tearoff=0)
        references_menu.add_command(label="Поставщик",
                                    command=lambda: self.show_table("SELECT * FROM postavchik", POSTAVCHIK_HEADERS))
        references_menu.add_command(label="Вид музукального инструмента",
                                    command=lambda: self.show_table("SELECT * FROM vid_muz_instr", VID_MUZ_INSTR_HEADERS))
        self.menu_bar.add_cascade(label="Справочники", menu=references_menu)

        # Меню "Таблицы"
        tables_menu = tk.Menu(self.menu_bar, tearoff=0)
        tables_menu.add_command(label="Муз инстр", command=lambda: self.show_table("SELECT * FROM muz_instr", MUZ_INSTR_HEADERS))
        tables_menu.add_command(label="Поступлние", command=lambda: self.show_table("SELECT * FROM postuplenie", POSTEUPLENIE_HEADERS))
        tables_menu.add_command(label="Продажи", command=lambda: self.show_table("SELECT * FROM sale", SALE_HEADERS))
        self.menu_bar.add_cascade(label="Таблицы", menu=tables_menu)

# Меню "Отчёты"
        reports_menu = tk.Menu(self.menu_bar, tearoff=0)
        reports_menu.add_command(label="Создать Отчёт", command=self.to_xlsx)
        self.menu_bar.add_cascade(label="Отчёты", menu=reports_menu)

        # Меню "Сервис"
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        help_menu.add_command(label="Руководство пользователя")
        help_menu.add_command(label="O программе")
        self.menu_bar.add_cascade(label="Сервис")

        
 # Настройка цветов меню
        file_menu.configure(bg='#555', fg='white')
        references_menu.configure(bg='#555', fg='white')
        tables_menu.configure(bg='#555', fg='white')
        reports_menu.configure(bg='#555', fg='white')
        help_menu.configure(bg='#555', fg='white')

        # Установка меню в главное окно
        self.config(menu=self.menu_bar)

        btn_width = 150
        pad = 5

        # Создание кнопок и виджетов для поиска и редактирования данных
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=0, column=1)
        ctk.CTkButton(btn_frame, text="добавить", width=btn_width, command=self.add).pack(pady=pad)
        ctk.CTkButton(btn_frame, text="удалить", width=btn_width, command=self.delete).pack(pady=pad)
        ctk.CTkButton(btn_frame, text="изменить", width=btn_width, command=self.change).pack(pady=pad)

    def search_in_table(self, table, search_terms, start_item=None):
        table.selection_remove(table.selection())  # Сброс предыдущего выделения

        items = table.get_children('')
        start_index = items.index(start_item) + 1 if start_item else 0

        for item in items[start_index:]:
            values = table.item(item, 'values')
            for term in search_terms:
                if any(term.lower() in str(value).lower() for value in values):
                    table.selection_add(item)
                    table.focus(item)
                    table.see(item)
                    return item  # Возвращаем найденный элемент

    def reset_search(self):
        if self.last_headers:
            self.table.selection_remove(self.table.selection())
        self.search_entry.delete(0, 'end')

    def search(self):
        if self.last_headers:
            self.current_item = self.search_in_table(self.table, self.search_entry.get().split(','))

    def search_next(self):
        if self.last_headers:
            if self.current_item:
                self.current_item = self.search_in_table(self.table, self.search_entry.get().split(','), start_item=self.current_item)
    
    
    def to_xlsx(self):
        if self.last_headers == POSTAVCHIK_HEADERS:
            sql_query = "SELECT * FROM postavchik"
            table_name = "postavchik"
        elif self.last_headers == VID_MUZ_INSTR_HEADERS:
            sql_query = "SELECT * FROM vid_muz_instr"
            table_name = "vid_muz_instr"
        elif self.last_headers == MUZ_INSTR_HEADERS:
            sql_query = "SELECT * FROM muz_instr"
            table_name = "muz_instr"
        elif self.last_headers == POSTEUPLENIE_HEADERS:
            sql_query = "SELECT * FROM postuplenie"
            table_name = "postuplenie"
        elif self.last_headers == SALE_HEADERS:
            sql_query = "SELECT * FROM sale"
            table_name = "sale"
        else: return

        dir = sys.path[0] + "\\export"
        os.makedirs(dir, exist_ok=True)
        path = dir + f"\\{table_name}.xlsx"

        # Подключение к базе данных SQLite
        conn = sqlite3.connect("res\\students_bd.db")
        cursor = conn.cursor()
        # Получите данные из базы данных
        cursor.execute(sql_query)
        data = cursor.fetchall()
        # Создайте DataFrame из данных
        df = pd.DataFrame(data, columns=self.last_headers)
        # Создайте объект writer для записи данных в Excel
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        # Запишите DataFrame в файл Excel
        df.to_excel(writer, 'Лист 1', index=False)
        # Сохраните результат
        writer.close()

        showinfo(title="Успешно", message=f"Данные экспортированы в {path}")

    def show_table(self, sql_query, headers = None):# Очистка фрейма перед отображением новых данных
        for widget in self.table_frame.winfo_children(): widget.destroy()

        # Подключение к базе данных SQLite
        conn = sqlite3.connect("res\\students_bd.db")
        cursor = conn.cursor()

        # Выполнение SQL-запроса
        cursor.execute(sql_query)
        self.last_sql_query = sql_query

        # Получение заголовков таблицы и данных
        if headers == None: # если заголовки не были переданы используем те что в БД
            table_headers = [description[0] for description in cursor.description]
        else: # иначе используем те что передали
            table_headers = headers
            self.last_headers = headers
        table_data = cursor.fetchall()

        # Закрытие соединения с базой данных
        conn.close()
            
        canvas = ctk.CTkCanvas(self.table_frame, width=865, height=480)
        canvas.pack(fill="both", expand=True)

        x_scrollbar = ttk.Scrollbar(self.table_frame, orient="horizontal", command=canvas.xview)
        x_scrollbar.pack(side="bottom", fill="x")

        canvas.configure(xscrollcommand=x_scrollbar.set)

        self.table = ttk.Treeview(self.table_frame, columns=table_headers, show="headings", height=23)
        for header in table_headers: 
            self.table.heading(header, text=header)
            self.table.column(header, width=len(header) * 10 + 100) # установка ширины столбца исходя длины его заголовка
        for row in table_data: self.table.insert("", "end", values=row)

        canvas.create_window((0, 0), window=self.table, anchor="nw")

        self.table.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))


    def update_table(self):
        self.show_table(self.last_sql_query, self.last_headers)

    def add(self):
        if self.last_headers == MUZ_INSTR_HEADERS:
            WindowBook("add")
        elif self.last_headers == POSTAVCHIK_HEADERS:
            WindowPostafshik("add")
        elif self.last_headers == POSTEUPLENIE_HEADERS:
            WindowStudent("add")
        elif self.last_headers == SALE_HEADERS:
            WindowFormulyar("add")
        elif self.last_headers == VID_MUZ_INSTR_HEADERS:
            WindowSchoolLibrary("add")
        else: return

    def delete(self):
        if self.last_headers:
            select_item = self.table.selection()
            if select_item:
                item_data = self.table.item(select_item[0])["values"]
            else:
                showerror(title="Ошибка", message="He выбранна запись")
                return
        else:
            return

        if self.last_headers == MUZ_INSTR_HEADERS:
            WindowBook("delete", item_data)
        elif self.last_headers == POSTAVCHIK_HEADERS:
            WindowPostafshik("delete", item_data)
        elif self.last_headers == POSTEUPLENIE_HEADERS:
            WindowStudent("delete", item_data)
        elif self.last_headers == SALE_HEADERS:
            WindowFormulyar("delete", item_data)
        elif self.last_headers == VID_MUZ_INSTR_HEADERS:
            WindowSchoolLibrary("delete", item_data)
        else: return
        
    def change(self):
        if self.last_headers:
            select_item = self.table.selection()
            if select_item:
                item_data = self.table.item(select_item[0])["values"]
            else:
                showerror(title="Ошибка", message="He выбранна запись")
                return
        else:
            return

        if self.last_headers == MUZ_INSTR_HEADERS:
            WindowBook("change", item_data)
        elif self.last_headers == POSTAVCHIK_HEADERS:
            WindowPostafshik("change", item_data)
        elif self.last_headers == POSTEUPLENIE_HEADERS:
            WindowStudent("change", item_data)
        elif self.last_headers == SALE_HEADERS:
            WindowFormulyar("change", item_data)
        elif self.last_headers == VID_MUZ_INSTR_HEADERS:
            WindowSchoolLibrary("change", item_data)
        else: return

class WindowBook(tk.Toplevel):
    def __init__(self, operation, select_row = None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())
        if select_row: self.select_row = select_row

        if operation == "add":
            tk.Label(self, text="Наименование").grid(row=0, column=0)
            self.name = tk.Entry(self, width=20)
            self.name.grid(row=0, column=1)

            tk.Label(self, text="ID Вида инструмента").grid(row=1, column=0)
            self.genre = tk.Entry(self, width=20)
            self.genre.grid(row=1, column=1)

            tk.Label(self, text="Цена").grid(row=2, column=0)
            self.author = tk.Entry(self, width=20)
            self.author.grid(row=2, column=1)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=5, column=0)
            tk.Button(self, text="Сохранить", command=self.add).grid(row=5, column=1, sticky="e")

        elif operation == "delete":
            tk.Label(self, text=f"Вы действиельно хотите удалить запись\nИз таблицы 'Муз инстр'?").grid(row=0, column=0, columnspan=2)
            tk.Label(self, text=f"Значение: {self.select_row[1]}").grid(row=1, column=0, columnspan=2)
            tk.Button(self, text="Да", command=self.delete, width=12).grid(row=2, column=0)
            tk.Button(self, text="Нет", command=self.quit_win, width=12).grid(row=2, column=1)
        
        elif operation == "change":
            tk.Label(self, text="Наименование поля").grid(row=0, column=0)
            tk.Label(self, text="Текушее значение ").grid(row=0, column=1)
            tk.Label(self, text="Новое значение   ").grid(row=0, column=2)

            tk.Label(self, text="Название").grid(row=1, column=0)
            tk.Label(self, text=self.select_row[1]).grid(row=1, column=1)
            self.name = tk.Entry(self, width=20)
            self.name.grid(row=1, column=2)

            tk.Label(self, text="ID Вида инструмента").grid(row=2, column=0)
            tk.Label(self, text=self.select_row[2]).grid(row=2, column=1)
            self.genre = tk.Entry(self, width=20)
            self.genre.grid(row=2, column=2)

            tk.Label(self, text="Цена").grid(row=3, column=0)
            tk.Label(self, text=self.select_row[3]).grid(row=3, column=1)
            self.author = tk.Entry(self, width=20)
            self.author.grid(row=3, column=2)
            
            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=6, column=0)
            tk.Button(self, text="Сохранить", command=self.change).grid(row=6, column=2, sticky="e")
    
    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()
    
    def add(self):
        name = self.name.get()
        genre = self.genre.get()
        author = self.author.get()
        
        if name and genre and author:
            try:
                conn = sqlite3.connect("res\\students_bd.db")
                cursor = conn.cursor()
                cursor.execute(f"INSERT INTO muz_instr (name, id_vid, price) VALUES (?, ?, ?)",
                            (name, genre, author))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM muz_instr WHERE id_muz_instr = ?", (self.select_row[0],))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        name = self.name.get() or self.select_row[1]
        genre = self.genre.get() or self.select_row[2]
        author = self.author.get() or self.select_row[3]
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute(f"UPDATE muz_instr SET (name, id_vid, price) = (?, ?, ?) WHERE id_muz_instr = {self.select_row[0]}",
                        (name, genre, author))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

class WindowPostafshik(tk.Toplevel):
    def __init__(self, operation, select_row = None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())
        if select_row: self.select_row = select_row

        if operation == "add":
            tk.Label(self, text="Название поставщика").grid(row=0, column=0)
            self.n_doc = tk.Entry(self, width=20)
            self.n_doc.grid(row=0, column=1)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=2, column=0)
            tk.Button(self, text="Сохранить", command=self.add).grid(row=2, column=1, sticky="e")

        elif operation == "delete":
            tk.Label(self, text=f"Вы действиельно хотите удалить запись\nИз таблицы 'Поставщики'?").grid(row=0, column=0, columnspan=2)
            tk.Label(self, text=f"Значение: {self.select_row[1]}").grid(row=1, column=0, columnspan=2)
            tk.Button(self, text="Да", command=self.delete, width=12).grid(row=2, column=0)
            tk.Button(self, text="Нет", command=self.quit_win, width=12).grid(row=2, column=1)
        
        elif operation == "change":
            tk.Label(self, text="Наименование поля").grid(row=0, column=0)
            tk.Label(self, text="Текушее значение ").grid(row=0, column=1)
            tk.Label(self, text="Новое значение   ").grid(row=0, column=2)

            tk.Label(self, text="№ Название поставщика").grid(row=1, column=0)
            tk.Label(self, text=self.select_row[1]).grid(row=1, column=1)
            self.n_doc = tk.Entry(self, width=20)
            self.n_doc.grid(row=1, column=2)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=3, column=0)
            tk.Button(self, text="Сохранить", command=self.change).grid(row=3, column=2, sticky="e")
    
    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()
    
    def add(self):
        n_doc = self.n_doc.get()
        if n_doc:
            try:
                conn = sqlite3.connect("res\\students_bd.db")
                cursor = conn.cursor()
                cursor.execute(f"INSERT INTO postavchik (name_postavchika) VALUES (?)", (n_doc))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM postavchik WHERE id_postavchika = ?", (self.select_row[0],))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        n_doc = self.n_doc.get() or self.select_row[1]
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute("UPDATE postavchik SET name_postavchika = ? WHERE id_postavchika = ?",
                           (n_doc, self.select_row[0]))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

class WindowStudent(tk.Toplevel):
    def __init__(self, operation, select_row = None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())
        if select_row: self.select_row = select_row

        if operation == "add":
            tk.Label(self, text="Дата").grid(row=0, column=0)
            self.fio = tk.Entry(self, width=20)
            self.fio.grid(row=0, column=1)

            tk.Label(self, text="ID Mуз инструмента").grid(row=1, column=0)
            self.gruop = tk.Entry(self, width=20)
            self.gruop.grid(row=1, column=1)

            tk.Label(self, text="Номер ТТН").grid(row=2, column=0)
            self.nomer_ttn = tk.Entry(self, width=20)
            self.nomer_ttn.grid(row=2, column=1)

            tk.Label(self, text="ID Поставщика").grid(row=3, column=0)
            self.id_postavchika = tk.Entry(self, width=20)
            self.id_postavchika.grid(row=3, column=1)

            tk.Label(self, text="Цена за 1 товар").grid(row=4, column=0)
            self.price_za_1_tov = tk.Entry(self, width=20)
            self.price_za_1_tov.grid(row=4, column=1)

            tk.Label(self, text="Количество").grid(row=5, column=0)
            self.count = tk.Entry(self, width=20)
            self.count.grid(row=5, column=1)

            tk.Label(self, text="Сумма").grid(row=6, column=0)
            self.suma = tk.Entry(self, width=20)
            self.suma.grid(row=6, column=1)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=7, column=0)
            tk.Button(self, text="Сохранить", command=self.add).grid(row=7, column=1, sticky="e")

        elif operation == "delete":
            tk.Label(self, text=f"Вы действиельно хотите удалить запись\nИз таблицы 'Книги'?").grid(row=0, column=0, columnspan=2)
            tk.Label(self, text=f"Значение: {self.select_row[1]}").grid(row=1, column=0, columnspan=2)
            tk.Button(self, text="Да", command=self.delete, width=12).grid(row=2, column=0)
            tk.Button(self, text="Нет", command=self.quit_win, width=12).grid(row=2, column=1)
        
        elif operation == "change":
            tk.Label(self, text="Наименование поля").grid(row=0, column=0)
            tk.Label(self, text="Текушее значение ").grid(row=0, column=1)
            tk.Label(self, text="Новое значение   ").grid(row=0, column=2)

            tk.Label(self, text="Дата").grid(row=0, column=0)
            self.fio = tk.Entry(self, width=20)
            self.fio.grid(row=0, column=1)

            tk.Label(self, text="ID Mуз инструмента").grid(row=1, column=0)
            self.gruop = tk.Entry(self, width=20)
            self.gruop.grid(row=1, column=1)

            tk.Label(self, text="Номер ТТН").grid(row=2, column=0)
            self.nomer_ttn = tk.Entry(self, width=20)
            self.nomer_ttn.grid(row=2, column=1)

            tk.Label(self, text="ID Поставщика").grid(row=3, column=0)
            self.id_postavchika = tk.Entry(self, width=20)
            self.id_postavchika.grid(row=3, column=1)

            tk.Label(self, text="Цена за 1 товар").grid(row=4, column=0)
            self.price_za_1_tov = tk.Entry(self, width=20)
            self.price_za_1_tov.grid(row=4, column=1)

            tk.Label(self, text="Количество").grid(row=5, column=0)
            self.count = tk.Entry(self, width=20)
            self.count.grid(row=5, column=1)

            tk.Label(self, text="Сумма").grid(row=6, column=0)
            self.suma = tk.Entry(self, width=20)
            self.suma.grid(row=6, column=1)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=10, column=0)
            tk.Button(self, text="Сохранить", command=self.change).grid(row=10, column=2, sticky="e")
    
    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()
    
    def add(self):
        fio = self.fio.get()
        gruop = self.gruop.get()
        nomer_ttn = self.nomer_ttn.get()
        id_postavchika = self.id_postavchika.get()
        price_za_1_tov = self.price_za_1_tov.get()
        count = self.count.get()
        suma = self.suma.get()
        if fio and gruop and nomer_ttn and id_postavchika and price_za_1_tov and count and suma:
            try:
                conn = sqlite3.connect("res\\students_bd.db")
                cursor = conn.cursor()
                cursor.execute(f"INSERT INTO postuplenie (date, id_muz_instr, nomer_ttn, id_postavchika, price_za_1_tov, count, sum) VALUES (?, ?, ?, ?, ?, ?,?)", (fio, gruop, nomer_ttn, id_postavchika, price_za_1_tov, count, suma))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM postuplenie WHERE id_postuplenie = ?", (self.select_row[0],))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        fio = self.fio.get() or self.select_row[1]
        gruop = self.gruop.get() or self.select_row[2]
        nomer_ttn = self.nomer_ttn.get() or self.select_row[3]
        id_postavchika = self.id_postavchika.get() or self.select_row[4]
        price_za_1_tov = self.price_za_1_tov.get() or self.select_row[5]
        count = self.count.get() or self.select_row[6]
        suma = self.suma.get() or self.select_row[7]
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute(f"UPDATE postuplenie SET (date, id_muz_instr, nomer_ttn, id_postavchika, price_za_1_tov, count, sum) = (?, ?, ?, ?, ?, ?,?) WHERE id_postuplenie = {self.select_row[0]}",
                        (fio, gruop, nomer_ttn, id_postavchika, price_za_1_tov, count, suma))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

class WindowFormulyar(tk.Toplevel):
    def __init__(self, operation, select_row = None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())
        if select_row: self.select_row = select_row

        if operation == "add":
            tk.Label(self, text="Дата").grid(row=0, column=0)
            self.date_vidach = tk.Entry(self, width=20)
            self.date_vidach.grid(row=0, column=1)

            tk.Label(self, text="ID Муз инстр").grid(row=1, column=0)
            self.date_vozvar = tk.Entry(self, width=20)
            self.date_vozvar.grid(row=1, column=1)

            tk.Label(self, text="Цена").grid(row=2, column=0)
            self.n_book = tk.Entry(self, width=20)
            self.n_book.grid(row=2, column=1)

            tk.Label(self, text="Количество").grid(row=3, column=0)
            self.n_student = tk.Entry(self, width=20)
            self.n_student.grid(row=3, column=1)

            tk.Label(self, text="Сумма").grid(row=4, column=0)
            self.summa = tk.Entry(self, width=20)
            self.summa.grid(row=4, column=1)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=5, column=0)
            tk.Button(self, text="Сохранить", command=self.add).grid(row=5, column=1, sticky="e")

        elif operation == "delete":
            tk.Label(self, text=f"Вы действиельно хотите удалить запись\nИз таблицы 'Книги'?").grid(row=0, column=0, columnspan=2)
            tk.Label(self, text=f"Значение: {self.select_row[1]}").grid(row=1, column=0, columnspan=2)
            tk.Button(self, text="Да", command=self.delete, width=12).grid(row=2, column=0)
            tk.Button(self, text="Нет", command=self.quit_win, width=12).grid(row=2, column=1)
        
        elif operation == "change":
            tk.Label(self, text="Наименование поля").grid(row=0, column=0)
            tk.Label(self, text="Текушее значение ").grid(row=0, column=1)
            tk.Label(self, text="Новое значение   ").grid(row=0, column=2)

            tk.Label(self, text="Дата").grid(row=1, column=0)
            tk.Label(self, text=self.select_row[1]).grid(row=1, column=1)
            self.date_vidach = tk.Entry(self, width=20)
            self.date_vidach.grid(row=1, column=2)

            tk.Label(self, text="ID Муз инстр").grid(row=2, column=0)
            tk.Label(self, text=self.select_row[2]).grid(row=2, column=1)
            self.date_vozvar = tk.Entry(self, width=20)
            self.date_vozvar.grid(row=2, column=2)

            tk.Label(self, text="Цена").grid(row=3, column=0)
            tk.Label(self, text=self.select_row[3]).grid(row=3, column=1)
            self.n_book = tk.Entry(self, width=20)
            self.n_book.grid(row=3, column=2)

            tk.Label(self, text="Количество").grid(row=4, column=0)
            tk.Label(self, text=self.select_row[4]).grid(row=4, column=1)
            self.n_student = tk.Entry(self, width=20)
            self.n_student.grid(row=4, column=2)

            tk.Label(self, text="Сумма").grid(row=5, column=0)
            tk.Label(self, text=self.select_row[5]).grid(row=5, column=1)
            self.summa = tk.Entry(self, width=20)
            self.summa.grid(row=5, column=2)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=6, column=0)
            tk.Button(self, text="Сохранить", command=self.change).grid(row=6, column=2, sticky="e")
    
    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()
    
    def add(self):
        date_vidach = self.date_vidach.get()
        date_vozvar = self.date_vozvar.get()
        n_book = self.n_book.get()
        n_student = self.n_student.get()
        summa = self.summa.get()
        if date_vidach and date_vozvar and n_book and n_student:
            try:
                conn = sqlite3.connect("res\\students_bd.db")
                cursor = conn.cursor()
                cursor.execute(f"INSERT INTO sale (date, id_muz_instr, price, kolichestvo, sum) VALUES (?, ?, ?, ?, ?)",
                            (date_vidach, date_vozvar, n_book, n_student, summa))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM sale WHERE id_sale = ?", (self.select_row[0],))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        date_vidach = self.date_vidach.get() or self.select_row[1]
        date_vozvar = self.date_vozvar.get() or self.select_row[2]
        n_book = self.n_book.get() or self.select_row[3]
        n_student = self.n_student.get() or self.select_row[4]
        summa = self.summa.get() or self.select.row[5]
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute(f'''
                           UPDATE sale SET (date, id_muz_instr, price, kolichestvo, sum) = (?, ?, ?, ?, ?) 
                           WHERE id_sale = {self.select_row[0]}''', (date_vidach, date_vozvar, n_book, n_student, summa))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

class WindowSchoolLibrary(tk.Toplevel):
    def __init__(self, operation, select_row = None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())
        if select_row: self.select_row = select_row

        if operation == "add":
            tk.Label(self, text="Вид муз инструмента").grid(row=0, column=0)
            self.name = tk.Entry(self, width=20)
            self.name.grid(row=0, column=1)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=5, column=0)
            tk.Button(self, text="Сохранить", command=self.add).grid(row=5, column=1, sticky="e")

        elif operation == "delete":
            tk.Label(self, text=f"Вы действиельно хотите удалить запись\nИз таблицы 'Книги'?").grid(row=0, column=0, columnspan=2)
            tk.Label(self, text=f"Значение: {self.select_row[1]}").grid(row=1, column=0, columnspan=2)
            tk.Button(self, text="Да", command=self.delete, width=12).grid(row=2, column=0)
            tk.Button(self, text="Нет", command=self.quit_win, width=12).grid(row=2, column=1)
        
        elif operation == "change":
            tk.Label(self, text="Наименование поля").grid(row=0, column=0)
            tk.Label(self, text="Текушее значение ").grid(row=0, column=1)
            tk.Label(self, text="Новое значение   ").grid(row=0, column=2)

            tk.Label(self, text="Вид муз инструмента").grid(row=1, column=0)
            tk.Label(self, text=self.select_row[1]).grid(row=1, column=1)
            self.name = tk.Entry(self, width=20)
            self.name.grid(row=1, column=2)

            tk.Button(self, text="Отмена", command=self.quit_win).grid(row=6, column=0)
            tk.Button(self, text="Сохранить", command=self.change).grid(row=6, column=2, sticky="e")
    
    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()
    
    def add(self):
        name = self.name.get()
        if name:
            try:
                conn = sqlite3.connect("res\\students_bd.db")
                cursor = conn.cursor()
                cursor.execute("INSERT INTO vid_muz_instr (name_vid_muz_instr) VALUES (?)", (name,))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM vid_muz_instr WHERE id_vid_mus_instr = ?", (self.select_row[0],))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        name = self.name.get() or self.select_row[1]
        try:
            conn = sqlite3.connect("res\\students_bd.db")
            cursor = conn.cursor()
            cursor.execute("UPDATE vid_muz_instr SET name_vid_muz_instr = ? WHERE id_vid_mus_instr = ?",
                           (name, self.select_row[0]))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))


if __name__ == "__main__":
    win = WindowMain()
    win.mainloop()
