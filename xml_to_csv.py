# -*- coding: Windows-1251 -*-
import xml.etree.ElementTree as ET
import csv
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


def countV():
    global c
    counter = i.find('ВсеВыплаты/Количество')

    if counter != None:
        counter = counter.text
    else:
        counter = ' '
    p.append(counter)
    c = int(counter)


def blank():
    blanks = ' '
    p.append(blanks)
    p.append(blanks)
    p.append(blanks)
    p.append(blanks)
    p.append(blanks)


def viplata(count):
    if count == max(kInt):
        for v in i.findall('.//Выплата'):
            addTag('ПризнакВыплаты', v, p)
            addTag('СуммаКвыплате', v, p)
            addTag('ДатаНачалаПериода', v, p)
            addTag('ДатаКонцаПериода', v, p)
            addTag('ВидВыплатыПоПЗ', v, p)
    elif count < max(kInt):
        for v in i.findall('.//Выплата'):
            addTag('ПризнакВыплаты', v, p)
            addTag('СуммаКвыплате', v, p)
            addTag('ДатаНачалаПериода', v, p)
            addTag('ДатаКонцаПериода', v, p)
            addTag('ВидВыплатыПоПЗ', v, p)
            for b in range(max(kInt) - count):
                 blank()


def header():
    global kInt
    k = []
    head = []
    for i in root.iter('ВсеВыплаты'):
        addTag('Количество', i, k)
    kInt = [int(i) for i in k]

    for i in headerF:
        head.append(i)
    for i in headerV * max(kInt):
        head.append(i)
    for i in headerS:
        head.append(i)

    return head


openFile()
tree = ET.parse(file)
root = tree.getroot()
output = open(file + '.csv', 'w', newline='', encoding='Windows-1251')
csvwriter = csv.writer(output, delimiter=';')
p = []
headerF = ['НомерВмассиве', 'НомерВыплатногоДела', 'КодРайона', 'СтраховойНомер', 'Фамилия', 'Имя', 'Отчество',
          'НомерСчета', 'КоличествоВыплат']
headerV = ['ПризнакВыплаты', 'СуммаКвыплате', 'ДатаНачалаПериода', 'ДатаКонцаПериода', 'ВидВыплатыПоПЗ']
headerS = ['ОтзываемаяСумма', 'ПризнакВозврата', 'СуммаВозвращена', 'КодВозврата', 'ПричинаПрекращенияВыплаты',
           'ДатаПрекращенияВыплаты', 'ДатаСмерти', 'НомерЗаписиАкта', 'ДатаЗаписиАкта', 'ДатаВыдачиДокумента']
csvwriter.writerow(header())


for i in root.iter('ОтчетПоПолучателю'):
    addTag('НомерВмассиве', i, p)
    addTag('НомерВыплатногоДела', i, p)
    addTag('КодРайона', i, p)
    addTag('СтраховойНомер', i, p)
    addTag('ФИО/Фамилия', i, p)
    addTag('ФИО/Имя', i, p)
    addTag('ФИО/Отчество', i, p)
    addTag('НомерСчета', i, p)
    countV()
    viplata(c)
    addTag('ОтзываемаяСумма', i, p)
    addTag('ПризнакВозврата', i, p)
    addTag('СуммаВозвращена', i, p)
    addTag('КодВозврата', i, p)
    addTag('ПричинаПрекращенияВыплаты', i, p)
    addTag('ДатаПрекращенияВыплаты', i, p)
    addTag('ДатаСмерти', i, p)
    addTag('НомерЗаписиАкта', i, p)
    addTag('ДатаЗаписиАкта', i, p)
    addTag('ДатаВыдачиДокумента', i, p)
    csvwriter.writerow(p)
    p.clear()


output.close()
f.close()

