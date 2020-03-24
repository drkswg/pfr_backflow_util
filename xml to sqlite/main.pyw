#!python3.6
# -*- coding: Windows-1251 -*-
import sys
sys.path.insert(0, 'pkgs')
import os
import re
import xml.etree.ElementTree as ET
import sqlite3
import csv
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
import design


class XMLBaseApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.xml_to_sql_run)
        self.pushButton_2.clicked.connect(self.sql_to_csv_run)
        self.pushButton_3.clicked.connect(self.xml_to_csv_run)

    def get_directory(self):
        global files

        files = []
        try:
            directory_first = QFileDialog.getExistingDirectory(self, 'Выберите папку', 'C:/')
        except:
            pass
        directory = os.listdir(directory_first)

        for i in directory:
            if re.search('ONVZ', i):
                files.append(directory_first + '/' + i)
            elif re.search('OVZR', i):
                files.append(directory_first + '/' + i)

    def addTag(self, tag, iterator, spisok):
        var = iterator.find(tag)
        if var != None:
            var = var.text
        else:
            var = ' '
        spisok.append(var)

    def xml_to_sql(self):
        p = []
        vip = []
        db = 'C:/Users/kodia/Desktop/xml to csv/xml to sqlite/test.db'
        conn = sqlite3.connect(db)
        cur = conn.cursor()

        for file in files:
            with open(file, 'r'):
                tree = ET.parse(file)
                root = tree.getroot()

                for i in root.iter('ОтчетПоПолучателю'):
                    self.addTag('НомерВыплатногоДела', i, p)
                    self.addTag('КодРайона', i, p)
                    self.addTag('СтраховойНомер', i, p)
                    self.addTag('ФИО/Фамилия', i, p)
                    self.addTag('ФИО/Имя', i, p)
                    self.addTag('ФИО/Отчество', i, p)
                    self.addTag('НомерСчета', i, p)
                    self.addTag('ВсеВыплаты/Количество', i, p)
                    for v in i.findall('.//Выплата'):
                        self.addTag('ПризнакВыплаты', v, vip)
                        self.addTag('СуммаКвыплате', v, vip)
                        self.addTag('ДатаНачалаПериода', v, vip)
                        self.addTag('ДатаКонцаПериода', v, vip)
                        self.addTag('ВидВыплатыПоПЗ', v, vip)
                        cur.execute("INSERT INTO payments VALUES (?,?,?,?,?,?)",
                                    (
                                        p[2], vip[0], vip[1], vip[2], vip[3], vip[4])
                                    )
                        vip.clear()
                    self.addTag('ОтзываемаяСумма', i, p)
                    self.addTag('СуммаВозвращена', i, p)
                    if re.search('ONVZ', file):
                        self.addTag('КодНевозврата', i, p)
                    elif re.search('OVZR', file):
                        self.addTag('КодВозврата', i, p)
                    self.addTag('ПричинаПрекращенияВыплаты', i, p)
                    self.addTag('ДатаПрекращенияВыплаты', i, p)
                    self.addTag('ДатаСмерти', i, p)
                    self.addTag('НомерЗаписиАкта', i, p)
                    self.addTag('ДатаЗаписиАкта', i, p)
                    self.addTag('ДатаВыдачиДокумента', i, p)
                    cur.execute("INSERT OR IGNORE INTO clients VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                (
                                    p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8],
                                    p[9], p[10], p[11], p[12], p[13], p[14], p[15], p[16])
                                )
                    p.clear()

        cur.execute(
            "DELETE FROM payments WHERE ROWID NOT IN("
            "SELECT MAX(ROWID) FROM payments GROUP BY СтраховойНомер, СуммаКвыплате)"
        )
        conn.commit()
        conn.close()
        QMessageBox.information(self, 'Результат', 'Файлы успешно загружены в БД')

    def xml_to_sql_run(self):
        self.get_directory()

        if len(files) > 0:
            self.xml_to_sql()
        else:
            QMessageBox.information(self, 'Результат', 'Отсутствуют файлы для загрузки')

    def npers(self):
        n = self.lineEdit.text()

        return n

    def if_exist(self):
        n = []
        n.append(self.npers())
        db = 'C:/Users/kodia/Desktop/xml to csv/xml to sqlite/test.db'
        conn = sqlite3.connect(db)
        cur = conn.cursor()

        client = cur.execute("SELECT * FROM clients WHERE СтраховойНомер=?;", n)
        client = list(client)

        conn.commit()
        conn.close()

        try:
            if client[0][2] == self.npers():
                QMessageBox.information(self, 'Результат', 'Данные по пенсионеру успешно выгружены ')
                return 0
        except:
            QMessageBox.information(self, 'Результат', 'Пенсионер отсутствует в базе')
            return 1

    def sql_to_csv(self):
        n = []
        n.append(self.npers())
        db = 'C:/Users/kodia/Desktop/xml to csv/xml to sqlite/test.db'
        output = 'C:/Users/kodia/Desktop/xml to csv/xml to sqlite/output.csv'
        conn = sqlite3.connect(db)
        cur = conn.cursor()

        with open(output, 'w', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(
                ['НомерВыплатногоДела', 'КодРайона', 'СтраховойНомер', 'Фамилия', 'Имя', 'Отчество',
                 'НомерСчета', 'КоличествоВыплат', 'ОтзываемаяСумма', 'СуммаВозвращена', 'КодВозврата',
                 'ПричинаПрекращенияВыплаты', 'ДатаПрекращенияВыплаты', 'ДатаСмерти', 'НомерЗаписиАкта',
                 'ДатаЗаписиАкта', 'ДатаВыдачиДокумента']
            )
            writer.writerows(cur.execute("SELECT * FROM clients WHERE СтраховойНомер=?;", n))
            writer.writerow(
                ['СтраховойНомер', 'ПризнакВыплаты', 'СуммаКвыплате',
                 'ДатаНачалаПериода', 'ДатаКонцаПериода', 'ВидВыплатыПоПЗ']
            )
            writer.writerows(cur.execute("SELECT * FROM payments WHERE СтраховойНомер=?;", n))

        conn.commit()
        conn.close()

    def sql_to_csv_run(self):
        if self.if_exist() == 0:
            self.sql_to_csv()
        else:
            pass

    def openFile(self):
        global file

        file = QFileDialog.getOpenFileName(self, 'Выбор файла', 'C:/', 'XML-файлы (*.xml)')[0]

    def check_file(self):
        if re.search('ONVZ', file):
            return 0
        elif re.search('OVZR', file):
            return 0
        else:
            return 1

    def countV(self, iterator, spisok):
        counter = iterator.find('ВсеВыплаты/Количество')

        if counter != None:
            counter = counter.text
        else:
            counter = ' '
        spisok.append(counter)
        c = int(counter)

        return c

    def viplata(self, count, iterator):
        if count == max(kInt):
            for v in iterator.findall('.//Выплата'):
                self.addTag('ПризнакВыплаты', v, p)
                self.addTag('СуммаКвыплате', v, p)
                self.addTag('ДатаНачалаПериода', v, p)
                self.addTag('ДатаКонцаПериода', v, p)
                self.addTag('ВидВыплатыПоПЗ', v, p)
        elif count < max(kInt):
            for v in iterator.findall('.//Выплата'):
                self.addTag('ПризнакВыплаты', v, p)
                self.addTag('СуммаКвыплате', v, p)
                self.addTag('ДатаНачалаПериода', v, p)
                self.addTag('ДатаКонцаПериода', v, p)
                self.addTag('ВидВыплатыПоПЗ', v, p)
                for b in range(max(kInt) - count):
                    self.blank(p)

    def blank(self, spisok):
        blanks = ' '
        spisok.append(blanks)
        spisok.append(blanks)
        spisok.append(blanks)
        spisok.append(blanks)
        spisok.append(blanks)

    def header(self, root):
        global kInt

        headerF = ['НомерВмассиве', 'НомерВыплатногоДела', 'КодРайона', 'СтраховойНомер', 'Фамилия', 'Имя', 'Отчество',
                   'НомерСчета', 'КоличествоВыплат']
        headerV = ['ПризнакВыплаты', 'СуммаКвыплате', 'ДатаНачалаПериода', 'ДатаКонцаПериода', 'ВидВыплатыПоПЗ']
        headerS = ['ОтзываемаяСумма', 'ПризнакВозврата', 'СуммаВозвращена', 'КодВозврата', 'ПричинаПрекращенияВыплаты',
                   'ДатаПрекращенияВыплаты', 'ДатаСмерти', 'НомерЗаписиАкта', 'ДатаЗаписиАкта', 'ДатаВыдачиДокумента']
        k = []
        head = []

        for i in root.iter('ВсеВыплаты'):
            self.addTag('Количество', i, k)
        kInt = [int(i) for i in k]

        for i in headerF:
            head.append(i)
        for i in headerV * max(kInt):
            head.append(i)
        for i in headerS:
            head.append(i)

        return head

    def xml_to_csv(self):
        global p
        p = []

        with open(file, 'r'):
            tree = ET.parse(file)
            root = tree.getroot()
            output = open(file + '.csv', 'w', newline='', encoding='Windows-1251')
            csvwriter = csv.writer(output, delimiter=';')
            csvwriter.writerow(self.header(root))

            for i in root.iter('ОтчетПоПолучателю'):
                self.addTag('НомерВмассиве', i, p)
                self.addTag('НомерВыплатногоДела', i, p)
                self.addTag('КодРайона', i, p)
                self.addTag('СтраховойНомер', i, p)
                self.addTag('ФИО/Фамилия', i, p)
                self.addTag('ФИО/Имя', i, p)
                self.addTag('ФИО/Отчество', i, p)
                self.addTag('НомерСчета', i, p)
                self.viplata(self.countV(i, p), i)
                self.addTag('ОтзываемаяСумма', i, p)
                self.addTag('ПризнакВозврата', i, p)
                self.addTag('СуммаВозвращена', i, p)
                if re.search('ONVZ', file):
                    self.addTag('КодНевозврата', i, p)
                elif re.search('OVZR', file):
                    self.addTag('КодВозврата', i, p)
                self.addTag('ПричинаПрекращенияВыплаты', i, p)
                self.addTag('ДатаПрекращенияВыплаты', i, p)
                self.addTag('ДатаСмерти', i, p)
                self.addTag('НомерЗаписиАкта', i, p)
                self.addTag('ДатаЗаписиАкта', i, p)
                self.addTag('ДатаВыдачиДокумента', i, p)
                csvwriter.writerow(p)
                p.clear()

            output.close()
            QMessageBox.information(self, 'Результат', 'Таблица успешно сформирована')

    def xml_to_csv_run(self):
        self.openFile()
        if self.check_file() == 0:
            self.xml_to_csv()
        elif self.check_file() == 1:
            QMessageBox.information(self, 'Результат', 'Неверный тип файла')


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = XMLBaseApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()