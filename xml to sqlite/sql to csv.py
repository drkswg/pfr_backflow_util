# -*- coding: Windows-1251 -*-
import sqlite3
import csv

db = 'C:/Users/kodia/Desktop/xml to csv/xml to sqlite/test.db'
output = 'C:/Users/kodia/Desktop/xml to csv/xml to sqlite/output.csv'
npers = input('Введите СНИЛС: ')
n = []
n.append(npers)


def export():
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

def ifexist():
    conn = sqlite3.connect(db)
    cur = conn.cursor()

    client = cur.execute("SELECT * FROM clients WHERE СтраховойНомер=?;", n)
    client = list(client)

    conn.commit()
    conn.close()

    try:
        if client[0][2] == npers:
            print('ok')
    except:
        print('ne ok')

ifexist()
export()

