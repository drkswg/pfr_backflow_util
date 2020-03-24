# -*- coding: Windows-1251 -*-
import xml.etree.ElementTree as ET
import sqlite3
from tkinter import filedialog as fd


def openFile():
    global file
    global f
    file = fd.askopenfilename()
    f = open(file)
    f.read()


def addTag(tag, iterator, spisok):
    var = iterator.find(tag)
    if var != None:
        var = var.text
    else:
        var = ' '
    spisok.append(var)


def parser():
    conn = sqlite3.connect(db)
    cur = conn.cursor()

    for i in root.iter('ОтчетПоПолучателю'):
        addTag('НомерВыплатногоДела', i, p)
        addTag('КодРайона', i, p)
        addTag('СтраховойНомер', i, p)
        addTag('ФИО/Фамилия', i, p)
        addTag('ФИО/Имя', i, p)
        addTag('ФИО/Отчество', i, p)
        addTag('НомерСчета', i, p)
        addTag('ВсеВыплаты/Количество', i, p)
        for v in i.findall('.//Выплата'):
            addTag('ПризнакВыплаты', v, vip)
            addTag('СуммаКвыплате', v, vip)
            addTag('ДатаНачалаПериода', v, vip)
            addTag('ДатаКонцаПериода', v, vip)
            addTag('ВидВыплатыПоПЗ', v, vip)
            cur.execute("INSERT INTO payments VALUES (?,?,?,?,?,?)",
                        (p[2], vip[0], vip[1], vip[2], vip[3], vip[4])
                        )
            vip.clear()
        addTag('ОтзываемаяСумма', i, p)
        addTag('СуммаВозвращена', i, p)
        if 'ONVZ' in file:
            addTag('КодНевозврата', i, p)
        elif 'OVZR' in file:
            addTag('КодВозврата', i, p)
        addTag('ПричинаПрекращенияВыплаты', i, p)
        addTag('ДатаПрекращенияВыплаты', i, p)
        addTag('ДатаСмерти', i, p)
        addTag('НомерЗаписиАкта', i, p)
        addTag('ДатаЗаписиАкта', i, p)
        addTag('ДатаВыдачиДокумента', i, p)
        cur.execute("INSERT INTO clients VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8],
                     p[9], p[10], p[11], p[12], p[13], p[14], p[15], p[16])
                    )
        p.clear()

    conn.commit()
    conn.close()


db = 'C:/Users/kodia/Desktop/xml to csv/xml to sqlite/test.db'

openFile()

tree = ET.parse(file)
root = tree.getroot()
p = []
vip = []

parser()





