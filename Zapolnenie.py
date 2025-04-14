from docxtpl import DocxTemplate

FIO=input("Введите ФИО:" )
Doljnost=input("Введите должность:" )
Kuda=input("Введите место назначение:" )
Zachem=input("Введите цель направления:" )
Data_bukv=input("Введите дату (Формат ЧИСЛО МЕСЯЦ ГОД: 01 января 2024):" )
Vremya_C=input("Введите время в месте назначения С (Формат ЧАС МИН: 11 13):" )
Vremya_Do=input("Введите время в месте назначения ДО (Формат ЧАС МИН: 11 13):" )

spisok_C=Vremya_C.split()
spisok_Do=Vremya_Do.split()

vremya=' '.join([spisok_C[0],"час.", spisok_C[1],"мин. до",spisok_Do[0],"час.",spisok_C[1],"мин."])


cifr=Data_bukv.split()


MBC={'января':"01",
        'февраля':"02",
        'марта':"03",
        'апреля':"04",
        'мая':"05",
        'июня':"06",
        'июля':"07",
        'августа':"08",
        'сентября':"09",
        'октября':"10",
        'ноября':"11",
        'декабря':"12"
}

spisok_bukv=Data_bukv.split()

spisok_cifr=spisok_bukv.copy()

spisok_cifr[1]=MBC[spisok_bukv[1]]
cifri='.'.join(spisok_cifr)

doc=DocxTemplate("test.docx")
context={'ФИО':FIO,
         'Должность':Doljnost,
         'Куда':Kuda,
         'Зачем':Zachem,
         'Дата_буквы':Data_bukv,
         'Дата_цифры':cifri,
         'Время': vremya
         
         }
doc.render(context)
doc.save("itog_test.docx")
