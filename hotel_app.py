import os
import sys
from os import getcwd, listdir
from os.path import join, split
import pandas as pd
import xlsxwriter
from PyQt5 import uic
from PyQt5.QtWidgets import (QMainWindow, QApplication)
import datetime
import xlrd
# from docx import Document


INTERFACE_DIRECTORY = join(split(getcwd())[0], "interface")
DATABASE_DIRECTORY = join(split(getcwd())[0], "database")
CHECK_IN_DIRECTORY = join(split(getcwd())[0], "check in documents")
CHECK_OUT_DIRECTORY = join(split(getcwd())[0], "check out documents")

NUMS = list(map(str, [1, 2, 3, 4, 5, 6, 7, 8, 9, 0]))

LOGIN_PARAMETER = "Логин"
PASSWORD_PARAMETER = "Пароль"
BAN_PARAMETER = "Бан"
REASON_PARAMETER = "Причина"


"""Конкурсное задание: профиль автоматизиция пбизнес-процесоов
Новицкий Лев"""

"""Dataformer-специальный класс для обработки excel таблиц и работы со временем"""


class DataFormer:
    def form(self, name):
        db_name = join(DATABASE_DIRECTORY, str(name) + ".xlsx")
        db = pd.read_excel(db_name)
        get_db = pd.concat([db])

        parameter_list = list(db)

        rows = list(get_db.shape)[0]

        data_list = []

        for i in range(len(parameter_list)):
            a = []
            for j in range(len(list(get_db.head(rows)[parameter_list[i]]))):
                a.append(str(list(get_db.head(rows)[parameter_list[i]])[j]))
            data_list.append(a)
        mass = []

        for i in range(len(data_list[0])):
            a = []
            for j in range(len(data_list)):
                a.append(str(data_list[j][i]))
            mass.append(a)
        return mass

    def first_row(self, name):
        db_name = join(DATABASE_DIRECTORY, str(name) + ".xlsx")
        db = pd.read_excel(db_name)
        get_db = pd.concat([db])

        parameter_list = list(get_db)
        return parameter_list

    def time_form(self):
        time = str(datetime.datetime.now())
        time = ("".join(time.split()[1]).split("."))[0]
        block_time = time
        return block_time

    def date_form(self):
        time = str(datetime.datetime.now())
        date = "".join(time.split()[0])
        return date

    def time_writer(self, name):
        name = str(name)
        wb = xlsxwriter.Workbook(join(DATABASE_DIRECTORY, f"{name}.xlsx"))
        wsh = wb.add_worksheet()
        fr = DataFormer().first_row(name)
        refill = DataFormer().form(name)
        time = str(datetime.datetime.now())
        refill.insert(0, fr)
        refill.append([time])
        print(refill)
        for i in range(len(refill)):
            for j in range(len(refill[i])):
                wsh.write(i, j, refill[i][j])

        wb.close()
        return refill


class Login(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(join(INTERFACE_DIRECTORY, "2nd_page_login_screen.ui"), self)
        self.cancel_b.clicked.connect(self.cancel)
        self.login_b.clicked.connect(self.login)
        self.admin_rb.toggled.connect(self.as_admin)
        self.manager_rb.toggled.connect(self.as_manager)

        self.failed_e = 0

        self.admin = False
        self.manager = False

    def as_admin(self):
        self.manager = False
        self.admin = True

    def as_manager(self):
        self.admin = False
        self.manager = True

    def get_time_dif(self):
        if len(DataFormer().form("system_block_time")) == 0:
            return []
        last_block = DataFormer().form("system_block_time")[-1][0]
        unblock_time = list(map(int, last_block.split(":")))
        second = unblock_time[2] + 180
        minute = unblock_time[1] + second // 60
        second %= 180
        hour = unblock_time[0] + minute // 60
        minute %= 180
        unblock_time = hour * 60 * 60 + minute * 60 + second

        now = list(map(int, DataFormer().time_form().split(":")))
        now = now[0] * 60 * 60 + now[1] * 60 + now[2]
        return [unblock_time, now]

    def check_time(self):
        if len(self.get_time_dif()) == 0:
            return True
        if self.get_time_dif()[0] <= self.get_time_dif()[1]:
            return True

        return False

    def login(self):
        if not self.check_time():
            time = list(map(int, self.get_time_dif()))
            minute = (time[0] - time[1]) // 60
            second = (time[0] - time[1]) % 60
            self.error_display.setText(f"Система заблокирована на {minute} минут, {second} секунд")
            return

        if not self.manager and not self.admin:
            self.error_display.setText("Вы не выбрали тип аккаунта")
            return

        if self.manager:
            if self.failed_e == 3:
                self.block_wb = xlsxwriter.Workbook(join(split(getcwd())[0], "database/system_block_time.xlsx"))
                self.block_wsh = self.block_wb.add_worksheet()

                fr = DataFormer().first_row("system_block_time")
                refill = DataFormer().form("system_block_time")
                refill.insert(0, fr)

                block = [DataFormer().time_form()]
                refill.append(block)

                for i in range(len(refill)):
                    for j in range(len(refill[i])):
                        self.block_wsh.write(i, j, refill[i][j])

                self.block_wb.close()
                return

            manager_db_name = join(split(getcwd())[0], "database/managers.xlsx")           #подключение excel файла через pandas
            manager_db = pd.read_excel(manager_db_name)
            manager_get_db = pd.concat([manager_db])

            rows = list(manager_get_db.shape)[0]

            manager_logins = list(manager_get_db.head(rows)[LOGIN_PARAMETER])              #создагние списков логинов и паролей из excel файла
            manager_passwords = list(manager_get_db.head(rows)[PASSWORD_PARAMETER])

            if self.login_input.text() == "" or self.password_input.text() == "":
                self.error_display.setText("Вы не указали логин или пароль")
                self.failed_e += 1
                DataFormer().time_writer("error_entry_time")
                return

            if self.login_input.text() not in manager_logins:
                self.error_display.setText("Управляющего с таким логином не существует")
                self.failed_e += 1
                DataFormer().time_writer("error_entry_time")
                return

            if manager_passwords[manager_logins.index(self.login_input.text())] != self.password_input.text():
                self.error_display.setText("Вы ввели неправильный пароль")
                self.failed_e += 1
                DataFormer().time_writer("error_entry_time")
                return

            self.open_man()

        if self.admin:
            if self.failed_e >= 3:
                self.block_wb = xlsxwriter.Workbook(join(split(getcwd())[0], "database/system_block_time.xlsx"))
                self.block_wsh = self.block_wb.add_worksheet()

                fr = DataFormer().first_row("system_block_time")
                refill = DataFormer().form("system_block_time")
                refill.insert(0, fr)

                block = [DataFormer().time_form()]
                refill.append(block)
                for i in range(len(refill)):
                    for j in range(len(refill[i])):
                        self.block_wsh.write(i, j, refill[i][j])

                self.block_wb.close()
                return

            admin_db_name = join(split(getcwd())[0], "database/admins.xlsx")  # подключение excel файла через pandas
            admin_db = pd.read_excel(admin_db_name)
            admin_get_db = pd.concat([admin_db])

            rows = list(admin_get_db.shape)[0]

            admin_logins = list(
                admin_get_db.head(rows)[LOGIN_PARAMETER])  # создагние списков логинов и паролей из excel файла
            admin_passwords = list(admin_get_db.head(rows)[PASSWORD_PARAMETER])
            admin_ban = list(admin_get_db.head(rows)[BAN_PARAMETER])
            admin_reason = list(admin_get_db.head(rows)[REASON_PARAMETER])
            self.hotel_num = list(admin_get_db.head(rows)["Гостиница"])

            if self.login_input.text() == "" or self.password_input.text() == "":
                self.error_display.setText("Вы не указали логин или пароль")
                self.failed_e += 1
                DataFormer().time_writer("error_entry_time")
                return

            if self.login_input.text() not in admin_logins:
                self.error_display.setText("Управляющего с таким логином не существует")
                self.failed_e += 1
                DataFormer().time_writer("error_entry_time")
                return

            if admin_ban[admin_logins.index(self.login_input.text())] == "да":
                self.error_display.setText(f"Вы заблокированы, причина: "
                                           f"{admin_reason[admin_logins.index(self.login_input.text())]}")
                return

            if admin_passwords[admin_logins.index(self.login_input.text())] != self.password_input.text():
                self.error_display.setText("Вы ввели неправильный пароль")
                self.failed_e += 1
                DataFormer().time_writer("error_entry_time")
                return

            with open(join(DATABASE_DIRECTORY, 'admin_entrance.txt'), 'w') as fout:
                print(self.hotel_num[admin_logins.index(self.login_input.text())], file=fout)

            with open(join(DATABASE_DIRECTORY, 'admin_pos.txt'), 'w') as fout:
                print(admin_logins.index(self.login_input.text()), file=fout)

            direct = listdir(DATABASE_DIRECTORY)
            name = "Hotel_" + str(self.hotel_num[0]) + ".xlsx"
            if name not in direct:
                self.error_display.setText("Гостиницы, к которой привязан администратор не существует")
                return

            self.open_adm()

    def cancel(self):
        self.login_input.clear()
        self.password_input.clear()

    def open_adm(self):
        self.close()
        self.adm = AdminCabinet()
        self.adm.show()

    def open_man(self):
        self.close()
        self.man = ManagerCabinet()
        self.man.show()


class AdminCabinet(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(join(INTERFACE_DIRECTORY, "3rd_page_admin.ui"), self)

        self.add_guest_b.clicked.connect(self.add_guest_blank)
        self.add_room_b.clicked.connect(self.add_room_blank)
        self.log_out_b.clicked.connect(self.log_out)

        self.find_b.clicked.connect(self.find)

        self.find_b.setAutoDefault(True)
        self.filter_input.returnPressed.connect(self.find_b.click)
        self.find_b.clicked.connect(self.guest_refresh)

        self.find_room_b.setAutoDefault(True)
        self.room_filter_input.returnPressed.connect(self.find_room_b.click)
        self.find_room_b.clicked.connect(self.guest_refresh)

        self.delete_rb.toggled.connect(self.delete_admin)
        self.delete_room_rb.toggled.connect(self.delete_room)

        self.delete_b.clicked.connect(self.delete)
        self.delete_room_b.clicked.connect(self.delete_room_process)

        self.find_room_b.clicked.connect(self.room_refresh)

        with open(join(DATABASE_DIRECTORY, 'admin_entrance.txt'), 'r') as fin:
            self.hotel_num = int(fin.read().split()[0])

        self.guest_mass = DataFormer().form("guests")
        for i in self.guest_mass:
            if i[-3] == self.hotel_num:
                self.guest_list.addItem(" ".join(i))

        self.room_mass = DataFormer().form("Hotel_" + str(self.hotel_num))
        for i in self.room_mass:
            self.room_list.addItem(" ".join(i))

        self.guest_wb = xlsxwriter.Workbook(join(DATABASE_DIRECTORY, "guests.xlsx"))
        self.guest_wsh = self.guest_wb.add_worksheet()

        self.room_wb = xlsxwriter.Workbook(join(DATABASE_DIRECTORY, f"Hotel_{self.hotel_num}.xlsx"))
        self.room_wsh = self.room_wb.add_worksheet()

    def add_guest_blank(self):
        self.gb = GuestBlank()
        self.gb.show()

    def add_room_blank(self):
        self.rb = RoomBlank()
        self.rb.show()

    def guest_refresh(self):
        self.guest_mass = DataFormer().form("guests")
        self.guest_list.clear()
        for i in self.guest_mass:
            self.guest_list.addItem(" ".join(i))

    def room_refresh(self):
        self.room_mass = DataFormer().form("Hotel_" + str(self.hotel_num))
        self.room_list.clear()
        for i in self.room_mass:
            self.room_list.addItem(" ".join(i))

    def log_out(self):
        self.close()
        self.log = Login()
        self.log.show()

    def delete_admin(self):
        self.delete = True

    def delete_room(self):
        self.delete_r = True

    def delete(self):
        if self.delete:
            pos = self.guest_list.currentRow()

            document = Document()
            document.add_heading('Документ о выселении', 0)
            guest = document.add_paragraph('Данные посетителя:')
            guest_data = document.add_paragraph('')
            for i in self.guest_mass[pos]:
                guest_data.add_run(" ")
                guest_data.add_run(str(i))

            admin = document.add_paragraph('Данные администратора:')
            admin_data = document.add_paragraph('')
            d = DataFormer().form("admins")
            with open(join(DATABASE_DIRECTORY, 'admin_pos.txt'), 'r') as fin:
                num = int(fin.read().split()[0])
            for i in d[num]:
                admin_data.add_run(" ")
                admin_data.add_run(str(i))

            doc_name = str(self.guest_mass[pos][0]) + " " + str(self.guest_mass[pos][1]) + " " + str(self.guest_mass[pos][2])
            document.save(join(CHECK_OUT_DIRECTORY, doc_name))

            hotel_view = xlsxwriter.Workbook(join(DATABASE_DIRECTORY, 'Hotel_' + str(self.hotel_num) + '.xlsx'))
            rooms_wsh = hotel_view.add_worksheet()

            name = "Hotel_" + str(self.hotel_num)
            rooms = DataFormer().form(name)

            print(self.guest_mass[pos][8], int(self.guest_mass[pos][8]) - 1, rooms[int(self.guest_mass[pos][8]) - 1], rooms[int(self.guest_mass[pos][8]) - 1][-1])

            rooms[int(self.guest_mass[pos][8]) - 1][-1] = "нет"
            fr = DataFormer().first_row('Hotel_' + str(self.hotel_num))
            rooms.insert(0, fr)
            for i in range(len(rooms)):
                for j in range(len(rooms[i])):
                    rooms_wsh.write(i, j, rooms[i][j])
            hotel_view.close()

            del self.guest_mass[pos]

            d = self.guest_mass
            refill = []
            refill.extend(d)
            fr = DataFormer().first_row("guests")
            refill.insert(0, fr)
            for i in range(len(refill)):
                for j in range(len(refill[i])):
                    self.guest_wsh.write(i, j, refill[i][j])

            self.guest_wb.close()

            self.guest_list.clear()
            for i in self.guest_mass:
                self.guest_list.addItem(" ".join(i))

    def delete_room_process(self):
        if self.delete_r:
            pos = self.room_list.currentRow()
            del self.room_mass[pos]

            d = self.room_mass
            refill = []
            refill.extend(d)
            fr = DataFormer().first_row("Hotel_" + str(self.hotel_num))
            refill.insert(0, fr)

            for i in range(len(refill)):
                print(refill[i])
                for j in range(len(refill[i])):
                    print(refill[i][j])
                    self.room_wsh.write(i, j, refill[i][j])

            self.room_wb.close()

            self.room_list.clear()
            for i in self.room_mass:
                self.guest_list.addItem(" ".join(i))

    def find(self):
        self.guest_list.clear()
        for i in self.guest_mass:
            if self.filter_input.text() in " ".join(i):
                self.guest_list.addItem(" ".join(i))

    def find_room(self):
        self.room_list.clear()
        for i in self.room_mass:
            if self.filter_input.text() in " ".join(i):
                self.room_list.addItem(" ".join(i))


class ManagerCabinet(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(join(INTERFACE_DIRECTORY, "3rd_page_manager.ui"), self)

        self.add_admin_b.clicked.connect(self.add_admin_blank)
        self.add_hotel_b.clicked.connect(self.add_hotel_blank)

        self.log_out_b.clicked.connect(self.log_out)

        self.find_admin_b.clicked.connect(self.find_admin)
        self.find_admin_b.setAutoDefault(True)
        self.admin_filter_input.returnPressed.connect(self.find_admin_b.click)
        self.find_admin_b.clicked.connect(self.admin_refresh)

        self.find_hotel_b.clicked.connect(self.find_hotel)
        self.find_hotel_b.setAutoDefault(True)
        self.hotel_filter_input.returnPressed.connect(self.find_hotel_b.click)
        self.find_hotel_b.clicked.connect(self.hotel_refresh)

        self.find_guest_b.clicked.connect(self.find_guest)
        self.find_guest_b.setAutoDefault(True)
        self.guest_filter_input.returnPressed.connect(self.find_guest_b.click)

        self.guest_mass = DataFormer().form("guests")
        for i in self.guest_mass:
            self.guest_list.addItem(" ".join(i))

        self.admin_mass = DataFormer().form("admins")
        for i in self.admin_mass:
            self.admin_list.addItem(" ".join(i))

        self.hotel_mass = DataFormer().form("hotels")
        for i in self.hotel_mass:
            self.hotel_list.addItem(" ".join(i))

        self.delete_admin_rb.toggled.connect(self.delete_admin)
        self.delete_hotel_rb.toggled.connect(self.delete_hotel)

        self.delete_admin_b.clicked.connect(self.delete_process_a)
        self.delete_hotel_b.clicked.connect(self.delete_process_h)

        self.delete_a = False
        self.delete_h = False
        self.delete_g = False

        self.ban_a = False

        self.ban_admin_b.clicked.connect(self.admin_ban_process)
        self.ban_admin_rb.toggled.connect(self.ban_admin)

        self.rooms_b.clicked.connect(self.view_rooms)
        self.show_rooms = 1

    def add_hotel_blank(self):
        self.hb = HotelBlank()
        self.hb.show()

    def add_admin_blank(self):
        self.ab = AdminBlank()
        self.ab.show()

    def log_out(self):
        self.close()
        self.log = Login()
        self.log.show()

    def delete_admin(self):
        self.delete_a = True

    def delete_hotel(self):
        self.delete_h = True

    def delete_guest(self):
        self.delete_g = True

    def ban_admin(self):
        self.ban_a = True

    def admin_ban_process(self):
        if self.ban_a:
            pos = self.admin_list.currentRow()
            with open(join(DATABASE_DIRECTORY, 'ban_line.txt'), 'w') as fout:
                print(pos, file=fout)

            self.ban = BanAdminPage()
            self.ban.show()

    def delete_process_a(self):
        if self.delete_a:
            self.admin_wb = xlsxwriter.Workbook(join(split(getcwd())[0], "database/admins.xlsx"))
            self.admin_wsh = self.admin_wb.add_worksheet()

            pos = self.admin_list.currentRow()
            del self.admin_mass[pos]

            d = self.admin_mass
            refill = []
            refill.extend(d)
            fr = DataFormer().first_row("admins")
            refill.insert(0, fr)
            for i in range(len(refill)):
                for j in range(len(refill[i])):
                    self.admin_wsh.write(i, j, refill[i][j])

            self.admin_wb.close()

            self.admin_list.clear()
            for i in self.admin_mass:
                self.admin_list.addItem(" ".join(i))

    def delete_process_h(self):
        if self.delete_h:
            self.hotel_wb = xlsxwriter.Workbook(join(split(getcwd())[0], "database/hotels.xlsx"))
            self.hotel_wsh = self.hotel_wb.add_worksheet()
            pos = self.hotel_list.currentRow()
            del self.hotel_mass[pos]

            d = self.hotel_mass
            refill = []
            refill.extend(d)
            fr = DataFormer().first_row("hotels")
            refill.insert(0, fr)
            for i in range(len(refill)):
                for j in range(len(refill[i])):
                    self.hotel_wsh.write(i, j, refill[i][j])

            self.hotel_wb.close()

            self.hotel_list.clear()
            for i in self.hotel_mass:
                self.hotel_list.addItem(" ".join(i))

    def find_admin(self):
        self.admin_list.clear()
        for i in self.admin_mass:
            if self.admin_filter_input.text() in " ".join(i):
                self.admin_list.addItem(" ".join(i))

    def find_hotel(self):
        self.hotel_list.clear()
        for i in self.hotel_mass:
            if self.hotel_filter_input.text() in " ".join(i):
                self.hotel_list.addItem(" ".join(i))

    def find_guest(self):
        self.guest_list.clear()
        for i in self.guest_mass:
            if self.guest_filter_input.text() in " ".join(i):
                self.guest_list.addItem(" ".join(i))

    def admin_refresh(self):
        self.admin_mass = DataFormer().form("admins")
        self.admin_list.clear()
        for i in self.admin_mass:
            self.admin_list.addItem(" ".join(i))

    def hotel_refresh(self):
        self.hotel_mass = DataFormer().form("hotels")
        self.hotel_list.clear()
        for i in self.hotel_mass:
            self.hotel_list.addItem(" ".join(i))

    def view_rooms(self):
        self.show_rooms += 1
        if self.show_rooms % 2 == 0:
            pos = self.hotel_list.currentRow()
            self.hotel_mass = DataFormer().form("hotels")
            hotel = self.hotel_mass[pos][0]
            self.hotel_list.clear()
            self.hotel_rooms = DataFormer().form("Hotel_" + str(hotel))
            for i in self.hotel_rooms:
                self.hotel_list.addItem(" ".join(i))
        else:
            self.hotel_list.clear()
            self.hotel_mass = DataFormer().form("hotels")
            for i in self.hotel_mass:
                self.hotel_list.addItem(" ".join(i))


class BanAdminPage(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(join(INTERFACE_DIRECTORY, "ban_admin_page.ui"), self)
        self.cancel_b.clicked.connect(self.cancel)
        self.ban_b.clicked.connect(self.ban)

    def cancel(self):
        self.close()

    def ban(self):
        self.adm_wb = xlsxwriter.Workbook(join(split(getcwd())[0], "database/admins.xlsx"))
        self.adm_wsh = self.adm_wb.add_worksheet()
        reason = self.ban_reason_input.toPlainText()
        with open(join(DATABASE_DIRECTORY, 'ban_line.txt'), 'r') as fin:
            pos = int(fin.read().split()[0])

        name = 'admins'
        fr = [DataFormer().first_row(name)]
        d = DataFormer().form(name)
        d[pos][-1] = reason
        d[pos][-2] = "да"
        refill = []
        refill.extend(fr)
        refill.extend(d)

        for i in range(len(refill)):
            for j in range(len(refill[i])):
                self.adm_wsh.write(i, j, refill[i][j])

        self.adm_wb.close()
        self.close()


class AdminBlank(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(join(INTERFACE_DIRECTORY, "add_admin.ui"), self)
        self.add.clicked.connect(self.add_admin)
        self.cancel_b.clicked.connect(self.cancel)

        self.inputs = [self.login_input, self.password_input, self.name_input,
                       self.surename_input, self.father_name_input,
                       self.phone_num_input, self.hotel_num_input]

        self.adm_wb = xlsxwriter.Workbook(join(split(getcwd())[0], "database/admins.xlsx"))
        self.adm_wsh = self.adm_wb.add_worksheet()

    def cancel(self):
        self.close()

    def add_admin(self):
        admin_data = [str(i.text()) for i in self.inputs]
        print(admin_data)
        hotels = DataFormer().form("hotels")
        print(hotels)
        fc = 0
        for i in range(len(hotels)):
            if hotels[i][0] != admin_data[-1]:
                fc += 1
        print(fc)
        if fc == len(hotels):
            self.error_display.setText("Отеля с таким номером не существует")
            return

        fr = DataFormer().first_row("admins")
        refill = DataFormer().form("admins")
        refill.append(admin_data)
        refill.insert(0, fr)
        for i in range(len(refill)):
            for j in range(len(refill[i])):
                self.adm_wsh.write(i, j, refill[i][j])

        self.adm_wb.close()
        self.close()


class GuestBlank(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(join(INTERFACE_DIRECTORY, "add_guest.ui"), self)

        self.add.clicked.connect(self.add_guest)
        self.cancel_b.clicked.connect(self.cancel)

        self.inputs = [self.name_input, self.surename_input,
                       self.father_name_input, self.birthday_input, self.gender_input,
                       self.phone_num_input,
                       self.passport_series_input, self.passport_number_input,
                       self.hotel_num_input, self.room_input, self.date_input]

        self.guest_wb = xlsxwriter.Workbook(join(split(getcwd())[0], "database/guests.xlsx"))
        self.guest_wsh = self.guest_wb.add_worksheet()

    def cancel(self):
        self.close()

    def add_guest(self):
        guest_data = [str(i.text()) for i in self.inputs]
        hotel = guest_data[8]
        files = os.listdir(DATABASE_DIRECTORY)
        name = "Hotel_" + str(hotel)
        file_name = "Hotel_" + str(hotel) + ".xlsx"
        rooms = DataFormer().form(name)

        if len(rooms) == 0:
            self.error_display.setText("Вы не можете заселить посетителя в этот номер, так как в гостинице не номеров")
            return

        for i in list(str(guest_data[8])):
            if i not in NUMS:
                self.error_display.setText("Номер гостиницы указан в неправильном формате")
                return

        for i in list(str(guest_data[9])):
            if i not in NUMS:
                self.error_display.setText("Номер в отеле указан в неправильном формате")
                return

        if file_name not in files:
            self.error_display.setText("Такого отеля не существует")
            return

        fc = 0
        for i in rooms:
            if i[0] != guest_data[8]:
                fc += 1
        if fc == len(rooms):
            self.error_display.setText("Такого номера не существует")
            return

        if rooms[int(guest_data[8]) - 1][-1] == "да":
            self.error_display.setText("Этот номер занят")
            return

        fr = DataFormer().first_row("guests")
        refill = DataFormer().form("guests")
        refill.append(guest_data)
        refill.insert(0, fr)

        for i in range(len(refill)):
            for j in range(len(refill[i])):
                self.guest_wsh.write(i, j, refill[i][j])

        self.guest_wb.close()

        document = Document()
        document.add_heading('Документ о заселении', 0)
        p = document.add_paragraph('Данные посетителя:')
        g = document.add_paragraph('')
        for i in guest_data:
            g.add_run(" ")
            g.add_run(str(i))

        d = DataFormer().form("admins")
        with open(join(DATABASE_DIRECTORY, 'admin_pos.txt'), 'r') as fin:
            self.admin_num = int(fin.read().split()[0])

        ad = document.add_paragraph('Данные падминистратора:')
        add = document.add_paragraph("")
        print(d, self.admin_num, sep="\n")
        for i in d[self.admin_num]:
            add.add_run(" ")
            add.add_run(str(i))

        doc_name = guest_data[0] + guest_data[1] + guest_data[2]
        document.save(join(CHECK_IN_DIRECTORY, doc_name))

        hotel_view = xlsxwriter.Workbook(join(DATABASE_DIRECTORY, 'Hotel_' + str(hotel) + '.xlsx'))
        rooms_wsh = hotel_view.add_worksheet()

        rooms[int(guest_data[8]) - 1][-1] = "да"
        fr = DataFormer().first_row('Hotel_' + str(hotel))
        rooms.insert(0, fr)
        for i in range(len(rooms)):
            for j in range(len(rooms[i])):
                rooms_wsh.write(i, j, rooms[i][j])
        hotel_view.close()
        self.close()


class HotelBlank(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(join(INTERFACE_DIRECTORY, "add_hotel.ui"), self)

        self.add.clicked.connect(self.add_hotel)
        self.cancel_b.clicked.connect(self.cancel)

        self.inputs = [self.hotel_input, self.levels_input, self.rooms_input,
                       self.country_input, self.city_input, self.street_input,
                       self.house_input]

        self.hotel_wb = xlsxwriter.Workbook(join(split(getcwd())[0], "database/hotels.xlsx"))
        self.hotel_wsh = self.hotel_wb.add_worksheet()

    def cancel(self):
        self.close()

    def add_hotel(self):
        hotel_data = [str(i.text()) for i in self.inputs]
        fr = DataFormer().first_row("hotels")

        for i in list(str(hotel_data[0])):
            if i not in NUMS:
                self.error_display.setText("Номер гостиницы указан в неправильном формате")
                return

        for i in list(str(hotel_data[1])):
            if i not in NUMS:
                self.error_display.setText("Кол-во этажей гостиницы указан в неправильном формате")
                return

        for i in list(str(hotel_data[1])):
            if i not in NUMS:
                self.error_display.setText("Кол-во номеров гостиницы указан в неправильном формате")
                return

        refill = DataFormer().form("hotels")
        refill.append(hotel_data)
        refill.insert(0, fr)

        for i in range(len(refill)):
            for j in range(len(refill[i])):
                self.hotel_wsh.write(i, j, refill[i][j])

        self.hotel_wb.close()

        hotel_view = xlsxwriter.Workbook(join(DATABASE_DIRECTORY, 'Hotel_' + str(hotel_data[0]) + '.xlsx'))
        rooms = hotel_view.add_worksheet()
        refill = [["Номер", "Кол-во комнат", "Площадь", "Занятость"]]
        for i in range(len(refill)):
            for j in range(len(refill[i])):
                rooms.write(i, j, refill[i][j])
        hotel_view.close()
        self.close()


class RoomBlank(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(join(INTERFACE_DIRECTORY, "add_room.ui"), self)
        self.inputs = [self.room_input, self.rooms_input, self.area_input]
        self.add_b.clicked.connect(self.add_room)
        self.cancel_b.clicked.connect(self.cancel)

    def cancel(self):
        self.close()

    def add_room(self):
        with open(join(DATABASE_DIRECTORY, 'admin_entrance.txt'), 'r') as fin:
            pos = int(fin.read().split()[0])

        self.room_wb = xlsxwriter.Workbook(join(DATABASE_DIRECTORY, f"Hotel_{str(pos)}.xlsx"))
        self.room_wsh = self.room_wb.add_worksheet()

        room_data = [str(i.text()) for i in self.inputs]

        for i in list(str(room_data[0])):
            if i not in NUMS:
                self.error_display.setText('Некорректный ввод номера')

        for i in list(str(room_data[1])):
            if i not in NUMS:
                self.error_display.setText('Некорректный ввод кол-ва комнат')

        for i in list(str(room_data[2])):
            if i not in NUMS:
                self.error_display.setText('Некорректный ввод площади номера')

        fr = DataFormer().first_row("Hotel_" + str(pos))
        d = DataFormer().form("Hotel_" + str(pos))

        refill = [i for i in d]
        refill.insert(0, fr)
        refill.append(room_data)

        for i in range(len(refill)):
            for j in range(len(refill[i])):
                self.room_wsh.write(i, j, refill[i][j])

        self.room_wb.close()
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = Login()
    ex.show()
    sys.exit(app.exec_())
