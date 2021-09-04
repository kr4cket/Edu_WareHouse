from tkinter import *
from tkinter import ttk
# project by Kr4ckeT
import time
import threading
import openpyxl
import random


class Color(object):
    def __init__(self, color, number, Type, syst):
        self.name = "Краска"
        self.color = color  # цвет краски
        self.Type = Type    # тип краски
        self.number = number  # количество краски
        if syst == "л":  # система счисления количества краски
            self.sys = "litr"  # литры
        else:
            self.sys = "mt"  # метры кубические

    def __str__(self):
        rep = self.name+self.color
        return rep

    def change_system(self):  # изменение системы счиления
        if self.sys == "litr":
            self.number = self.number / 1000
            self.syst = "м^3"
            self.sys = "mt"
        else:
            self.number = int(self.number * 1000)
            self.syst = "л"
            self.sys = "litr"

    def check_rastv(self):  # проверка на необходимость раствора
        if self.Type == "Нитроэмаль":
            return "Да"
        else:
            return "Нет"

    # функции для вывода атрибутов класса в таблицы
    def name(self):
        rep = self.name
        return rep

    def clr(self):
        rep = self.color
        return rep

    def tp(self):
        rep = self.Type
        return rep

    def num(self):
        rep = self.number
        return rep

    def system(self):
        if self.sys == "litr":
            syst = "л"
        else:
            syst = "м3"
        return syst

    # функция добавления числа на склад
    def add_number(self, numb):
        self.number += numb

    # функция получения товара со склада
    def remove_number(self, numb):
        self.number -= numb


class Solvent(object):
    def __init__(self, nazv, number, syst):
        self.name = "Растворитель"
        self.nazv = nazv  # название растворителя
        self.number = number  # количество в ед. измерения
        if syst == "л":
            self.sys = "litr"
        else:
            self.sys = "mt"

    def __str__(self):
        rep = self.name + self.nazv
        return rep

    def change_system(self):  # изменение системы счиления
        if self.sys == "litr":
            self.number = self.number / 1000
            self.syst = "м^3"
            self.sys = "mt"
        else:
            self.number = int(self.number * 1000)
            self.syst = "л"
            self.sys = "litr"

    # функции для вывода атрибутов класса

    def name(self):
        rep = self.name
        return rep

    def nzv(self):
        rep = self.nazv
        return rep

    def num(self):
        rep = self.number
        return rep

    def system(self):
        if self.sys == "litr":
            syst = "л"
        else:
            syst = "м3"
        return syst
    # функция добавления числа на склад
    def add_number(self, numb):
        self.number += numb

    # функция получения товара со склада
    def remove_number(self, numb):
        self.number -= numb


class Nail(object):
    def __init__(self, nazv, number, syst, lenght):
        self.name = "Гвозди"
        self.nazv = nazv # название гвоздей
        self.number = number # количество гвоздей в ед.изм
        if syst == "шт":
            self.sys = "st"
        else:
            self.sys = "kg"

        self.lenght = lenght

    def __str__(self):
        rep = self.name + self.nazv + "-" + str(self.lenght)
        return rep

    # функции для вывода атрибутов класса
    def name(self):
        rep = self.name
        return rep

    def nzv(self):
        rep = self.nazv
        return rep

    def num(self):
        rep = self.number
        return rep

    def system(self):
        if self.sys == "st":
            syst = "шт"
        else:
            syst = "кг"
        return syst

    def lenght(self):
        rep = self.lenght
        return rep

    def change_system(self): # изменение системы счиления
        if self.sys == "st":
            self.number = self.number / 100
            self.sys = "kg"
        else:
            self.number = int(self.number * 100)
            self.sys = "st"

        # функция добавления числа на склад

    def add_number(self, numb):
        self.number += numb

        # функция получения товара со склада

    def remove_number(self, numb):
        self.number -= numb


class Application(Frame):
    def __init__(self, master):
        super(Application, self).__init__(master)
        self.grid()
        self.create_widgets()

    def create_widgets(self):
        # Открытие файла
        self.open_file()
        # Для заполнения пространства невидимым
        Label(self,
              text=""
              ).grid(row=4, column=0, sticky='W')
        Label(self,
              text="           "
              ).grid(row=5, column=2, sticky='W')
        ##
        Label(self,
              text="Заказать товар со склада"
              ).grid(row=0, column=0, sticky='N', columnspan=2)
        Label(self,
              text="Тип товара"
              ).grid(row=1, column=0, sticky='W')
        Label(self,
              text="Название"
              ).grid(row=2, column=0, sticky='W')
        Label(self,
              text="Количество"
              ).grid(row=3, column=0, sticky='W')

        self.Name = Entry(self)
        self.Name.grid(row=1, column=1, sticky='W')
        self.Nazv = Entry(self)
        self.Nazv.grid(row=2, column=1, sticky='W')
        self.Kol = Entry(self)
        self.Kol.grid(row=3, column=1, sticky='W')
        Button(self,
               text="         Заказать         ",
               command=self.give_item
               ).grid(row=5, column=0, sticky='N')
        Button(self,
               text="         Получить          ",
               command=self.get_item
               ).grid(row=5, column=1, sticky='W')
        # Таблица
        self.Table = ttk.Treeview(self, show ='headings', selectmode = "extended")
        self.Table.grid(row=7, column=0, columnspan=4)
        scroll_pane = ttk.Scrollbar(self, command=self.Table.yview)
        scroll_pane.grid(row=7, column=5, sticky=E, ipady = 87)
        self.Table.configure(yscrollcommand=scroll_pane.set)
        # текстовое поле для вывода результата
        self.table = Text(self, width=30, height=1, wrap=WORD)
        self.table.grid(row=5, column=2, columnspan=2, sticky="E")
        # Отметки для выбора таблицы
        self.table_choose = StringVar()
        self.table_choose.set(None)
        self.value = ["all", "color", "solvent", "nail"]
        Radiobutton(self,
                    text="Весь склад",
                    variable=self.table_choose,
                    value=self.value[0],
                    command=self.read_file
                    ).grid(row=1, column=3, sticky="E")
        Radiobutton(self,
                    text="Краски",
                    variable=self.table_choose,
                    value=self.value[1],
                    command=self.read_file
                    ).grid(row=2, column=3, sticky="E")
        Radiobutton(self,
                    text="Растворители",
                    variable=self.table_choose,
                    value=self.value[2],
                    command=self.read_file
                    ).grid(row=3, column=3, sticky="E")
        Radiobutton(self,
                    text="Гвозди",
                    variable=self.table_choose,
                    value=self.value[3],
                    command=self.read_file
                    ).grid(row=4, column=3, sticky="E")
        ##
        Label(self,
              text="Выбор таблицы:"
              ).grid(row=0, column=3, sticky='E')
        Button(self,
               text="Изменить  ед.изм",
               command=self.change_sys
               ).grid(row=1, column=2, sticky='N')
        Button(self,
               text=" Сохранить файл ",
               command=self.save_file
               ).grid(row=2, column=2, sticky='N')
        self.read_file()

    def open_file(self):
        self.sklad = []
        cl = 2
        nl = 2
        slv = 2
        self.workbook = openpyxl.open("test.xlsx")
        self.sheet = self.workbook.worksheets[0]
        for row in range(2, self.sheet.max_row+1):
            if self.sheet['A' + str(row)].value == "Краска":
                self.sheet = self.workbook.worksheets[1]
                n = Color(self.sheet['A' + str(cl)].value, self.sheet['B' + str(cl)].value,
                          self.sheet['C' + str(cl)].value, self.sheet['D' + str(cl)].value)
                self.sklad.append(n)
                cl += 1
                self.sheet = self.workbook.worksheets[0]
            elif self.sheet['A' + str(row)].value == "Растворитель":
                self.sheet = self.workbook.worksheets[2]
                n = Solvent(self.sheet['A' + str(slv)].value, self.sheet['B' + str(slv)].value,
                            self.sheet['C' + str(slv)].value)
                self.sklad.append(n)
                slv += 1
                self.sheet = self.workbook.worksheets[0]
            else:
                self.sheet = self.workbook.worksheets[3]
                n = Nail(self.sheet['A' + str(nl)].value, self.sheet['B' + str(nl)].value,
                         self.sheet['C' + str(nl)].value, self.sheet['D' + str(nl)].value)
                self.sklad.append(n)
                nl += 1
                self.sheet = self.workbook.worksheets[0]

    def read_file(self):
        try:
            v = self.table_choose.get()
            v = self.value.index(v)
        except:
            v = 0
        if v == 0:
            for item in self.Table.get_children():
                self.Table.delete(item)
            HEADS = ["Тип товара", "Название", "Количество"]
            self.Table['columns'] = HEADS
            for head in HEADS:
                self.Table.heading(head, text=head, anchor="center")
                self.Table.column(head, minwidth=0, width=160, stretch=NO, anchor="center")
            self.Table.column("Количество", width=160, stretch=True, minwidth=0)
            self.Table.event_generate("<<ThemeChanged>>")
            for i in self.sklad:
                j = i.name
                if j == "Краска":
                    n = i.clr()
                    nu = str(i.num()) + " " + i.system()
                    row = (j, n, nu)
                    self.Table.insert('', END, values=row)
                elif j == "Растворитель":
                    n = i.nazv
                    nu = str(i.num()) + " " + i.system()
                    row = (j, n, nu)
                    self.Table.insert('', END, values = row)
                elif j == "Гвозди":
                    n = i.nazv
                    nu = str(i.num()) + " " + i.system()
                    row = (j, n, nu)
                    self.Table.insert('', END, values=row)
        elif v == 1:
            for item in self.Table.get_children():
                self.Table.delete(item)
            HEADS = ["Название", "Тип краски", "Количество", "Необходимость растворителя"]
            self.Table['columns'] = HEADS
            for head in HEADS:
                self.Table.heading(head, text=head, anchor="center")
                self.Table.column(head, minwidth=0, width=100, stretch=NO, anchor="center")
                self.Table.column("Необходимость растворителя", width=180, stretch=True, minwidth=0)
                self.Table.event_generate("<<ThemeChanged>>")
            for i in self.sklad:
                j = i.name
                if j == "Краска":
                    n = i.clr()
                    nu = str(i.num()) + i.system()
                    t = i.tp()
                    neob = i.check_rastv()
                    row_1 = (n, t, nu, neob)
                    self.Table.insert('', END, values=row_1)
        elif v == 2:
            for item in self.Table.get_children():
                self.Table.delete(item)
            HEADS = ["Название", "Количество"]
            self.Table['columns'] = HEADS
            for head in HEADS:
                self.Table.heading(head, text=head, anchor="center")
                self.Table.column(head, minwidth=0, width=240, stretch=NO, anchor="center")
            self.Table.column("Количество", width=240, stretch=True, minwidth=0)
            self.Table.event_generate("<<ThemeChanged>>")
            for i in self.sklad:
                j = i.name
                if j == "Растворитель":
                    n = i.nazv
                    nu = str(i.num()) + i.system()
                    row_2 = (n, nu)
                    self.Table.insert('', END, values=row_2)
        elif v == 3:
            for item in self.Table.get_children():
                self.Table.delete(item)
            HEADS = ["Название", "Количество", "Длина(в мм)"]
            self.Table['columns'] = HEADS
            for head in HEADS:
                self.Table.heading(head, text=head, anchor="center")
                self.Table.column(head, minwidth=0, width=160, stretch=NO, anchor="center")
            self.Table.column("Длина(в мм)", width=160, stretch=True, minwidth=0)
            self.Table.event_generate("<<ThemeChanged>>")
            for i in self.sklad:
                j = i.name
                if j == "Гвозди":
                    n = i.nazv
                    nu = str(i.num()) + " " + i.system()
                    l = str(i.lenght)
                    row_3 = (n,nu,l)
                    self.Table.insert('', END, values=row_3)
        text = "Таблица " + str(v+1) + " выведена на экран!"
        self.table.config(state=NORMAL)
        self.table.delete(0.0, END)
        self.table.insert(0.0, text)
        self.table.config(state=DISABLED)

    def save_file(self):
        row = 2
        c = 2
        d = 2
        e = 2
        for item in self.sklad:
            if item.name == "Краска":
                self.sheet = self.workbook.worksheets[0]
                self.sheet[row][0].value = item.name
                self.sheet[row][1].value = item.color
                self.sheet[row][2].value = item.number
                self.sheet = self.workbook.worksheets[1]
                self.sheet[c][0].value = item.color
                self.sheet[c][1].value = item.number
                self.sheet[c][2].value = item.Type
                self.sheet[c][3].value = item.system()
                self.sheet = self.workbook.worksheets[0]
                c += 1
                row += 1
            elif item.name == "Растворитель":
                self.sheet = self.workbook.worksheets[0]
                self.sheet[row][0].value = item.name
                self.sheet[row][1].value = item.nazv
                self.sheet[row][2].value = item.number
                self.sheet = self.workbook.worksheets[2]
                self.sheet[d][0].value = item.nazv
                self.sheet[d][1].value = item.number
                self.sheet[d][2].value = item.system()
                self.sheet = self.workbook.worksheets[0]
                d += 1
                row += 1
            elif item.name == "Гвозди":
                self.sheet = self.workbook.worksheets[0]
                self.sheet[row][0].value = item.name
                self.sheet[row][1].value = item.nazv
                self.sheet[row][2].value = item.number
                self.sheet = self.workbook.worksheets[3]
                self.sheet[e][0].value = item.nazv
                self.sheet[e][1].value = item.number
                self.sheet[e][2].value = item.system()
                self.sheet[e][3].value = item.lenght
                self.sheet = self.workbook.worksheets[0]
                e += 1
                row += 1
        text = "Данные успешно сохранены!"
        self.table.config(state=NORMAL)
        self.table.delete(0.0, END)
        self.table.insert(0.0, text)
        self.table.config(state=DISABLED)
        self.workbook.save("test.xlsx")
        self.workbook.close()

    def change_sys(self):
        try:
            v = self.table_choose.get()
            v = self.value.index(v)
        except:
            v = 0
        if v == 0:
            for i in self.sklad:
                i.change_system()
            self.read_file()
        elif v == 1:
            for i in self.sklad:
                j = i.name
                if j == "Краска":
                    i.change_system()
                self.read_file()
        elif v == 2:
            for i in self.sklad:
                j = i.name
                if j == "Растворитель":
                    i.change_system()
                self.read_file()
        elif v == 3:
            for i in self.sklad:
                j = i.name
                if j == "Гвозди":
                    i.change_system()
                self.read_file()

    def get_item(self):
        v = None
        it = self.Name.get() + self.Nazv.get()
        for i in self.sklad:
            if self.Name.get() and i.name == "Краска":
                if it == (i.name + i.color):
                    v = i
            elif self.Name.get() and i.name == "Растворитель":
                if it == (i.name + i.nazv):
                    v = i
            elif self.Name.get() and i.name == "Гвозди":
                if it == (i.name + i.nazv + "-" + str(i.lenght)):
                    v = i
        if v :
            if int(self.Kol.get()) > v.number:
                text = "Такого количества товара нет на складе!"
                self.table.config(state=NORMAL)
                self.table.delete(0.0, END)
                self.table.insert(0.0, text)
                self.table.config(state=DISABLED)
            else:
                v.remove_number(int(self.Kol.get()))
                self.read_file()
                text = "Товар получен со склада!"
                self.table.config(state=NORMAL)
                self.table.delete(0.0, END)
                self.table.insert(0.0, text)
                self.table.config(state=DISABLED)
        else:
            text = "Такого товара нет на складе!"
            self.table.config(state=NORMAL)
            self.table.delete(0.0, END)
            self.table.insert(0.0, text)
            self.table.config(state=DISABLED)
    def give_item(self):
        v = None
        self.t = 0

        def sleep(i):
            if v:
                    text = "Товар заказан!"
                    self.table.config(state=NORMAL)
                    self.table.delete(0.0, END)
                    self.table.insert(0.0, text)
                    self.table.config(state=DISABLED)
                    time.sleep(random.randrange(5, 20) + 1)
                    v.add_number(int(self.Kol.get()))
                    self.read_file()
            else:
                text = "Такого товара нет на складе!"
                self.table.config(state=NORMAL)
                self.table.delete(0.0, END)
                self.table.insert(0.0, text)
                self.table.config(state=DISABLED)
            self.t -= 1
        it = self.Name.get() + self.Nazv.get()
        for i in self.sklad:
            if self.Name.get() and i.name == "Краска":
                if it == (i.name + i.color):
                    v = i
            elif self.Name.get() and i.name == "Растворитель":
                if it == (i.name + i.nazv):
                    v = i
            elif self.Name.get() and i.name == "Гвозди":
                if it == (i.name + i.nazv + "-" + str(i.lenght)):
                    v = i
        th = threading.Thread(target=sleep, args=(self.t,))
        self.t += 1
        th.start()


root = Tk()
root.resizable(width=False, height=False)
root.title("Склад")
root.geometry("500x375")
app = Application(root)
root.mainloop()
