import requests
import openpyexcel
from bs4 import BeautifulSoup
from xlsx2csv import Xlsx2csv
import pandas as pd
import time
import os
import  csv

#import csv


if(os.path.exists("ID.txt")):

    print("Введите время задержки между запросами в секундах")
    time_sleep = input()
    while(time_sleep.isdigit()!=True):
        print("Введите время задержки между запросами в секундах")
        time_sleep = input()

    time_sleep = int(time_sleep)
    if(time_sleep<0): time_sleep=0



    ids = open("ID.txt", encoding='utf-8').readlines()
    print("Всего объектов:" + str(len(ids)))
    try:
        wb = openpyexcel.load_workbook(filename='Итог.xlsx')
    except:
        wb = openpyexcel.Workbook()

    try:
        os.remove('Итог.csv')
    except:
        None

    csv_ = open('Итог.csv', 'w', encoding='cp1251')


    ws = wb.active

    for row in ws:
        for el in row:
            el.value=""

    ws.cell( 1, 1).value = "URL 1"
    ws.cell( 1, 2).value = "URL 2"
    ws.cell( 1, 3).value = "Название"
    ws.cell( 1, 4).value = "Адрес"
    ws.cell( 1, 5).value = "Метро"
    ws.cell( 1, 6).value = "Класс жилья"
    ws.cell( 1, 7).value = "Квартиры от"
    ws.cell( 1, 8).value = "Квартиры до"
    ws.cell( 1, 9).value = "Цена от"
    ws.cell( 1, 10).value = "Цена до"
    ws.cell( 1, 11).value = "Срок сдачи"
    ws.cell( 1, 12).value = "Тип договора"
    ws.cell( 1, 13).value = "Очереди"
    ws.cell( 1, 14).value = "Число корпусов"
    ws.cell( 1, 15).value = "Этажность"
    ws.cell( 1, 16).value = "Число квартир"
    ws.cell( 1, 17).value = "Высота потолков"
    ws.cell( 1, 18).value = "Тип дома"
    ws.cell( 1, 19).value = "Отделка"
    ws.cell( 1, 20).value = "Машиноместа в паркинге"
    ws.cell( 1, 21).value = "Дополнительные преимущества"
    ws.cell( 1, 22).value = "Фото"

    """csv_.write('URL 1;'
               'URL 2;'
               'Название;'
               'Адрес;Метро;'
               'Класс жилья;'
               'Квартиры от;'
               'Квартиры до;'
               'Цена от;'
               'Цена до;'
               'Срок сдачи;'
               'Тип договора;'
               'Очереди;'
               'Число корпусов;'
               'Этажность;'
               'Число квартир;'
               'Высота потолков;'
               'Тип дома;Отделка;'
               'Машиноместа в паркинге;'
               'Дополнительные преимущества;'
               'Фото'
               '\n')"""

    id_now = 1

    for id in ids:

        try:

            id = id[0:len(id) - 1]

            url = ("https://realty.yandex.ru/newbuilding/" + str(id))

            r = requests.get(url, allow_redirects = False)

            soup = BeautifulSoup(r.content, "html.parser")

            url2 = r.headers.get('location')

            r = requests.get(url2, allow_redirects=False)
            '''allow_redirects = False'''

            soup = BeautifulSoup(r.content, "html.parser")


            name = soup.find('h1', {'class': 'SiteCardHeader__title'}).string

            address = soup.find('div', {'class': 'SiteCardHeader__address'}).string

            metro_all = soup.find_all('span', {'class': 'MetroStation__title'})
            metro =""
            for m in metro_all:
                metro+=m.string+','

            if len(metro) >1:
                metro = metro[:-1]

            Info =  soup.find_all('div', {'class': 'SiteCardInfo__features-item'})

            class_ = ""
            area_from = ""
            area_to = ""
            price_from = ""
            price_to = ""
            deadline = ""

            for i in Info:

                if i.contents[0].string == "Класс жилья":
                    class_ = i.contents[1].string

                elif i.contents[0].string == "Квартиры":
                    area_from = i.contents[1].string.split(" до")[0].replace("от ","")
                    area_to = i.contents[1].string.split(" до ")[1].replace(" м²","")

                elif i.contents[0].string == "Цена":
                    price_from = i.contents[1].string.split(" до")[0].replace("от ", "")
                    price_to = i.contents[1].string.split(" до ")[1].replace(" ₽", "")

                elif i.contents[0].string == "Срок сдачи":
                    deadline = i.contents[1].string

            Info2 = soup.find_all('div', {'class': 'CardFeatures__itemBody'})

            type_of_contract = ""
            queues = ""
            number_of_buildings = ""
            storeys = ""
            number_of_apartments = ""
            ceiling_height = ""
            house_type = ""
            finish = ""
            parking = ""


            for i in Info2:

                if i.contents[0].string == "Тип договора":
                    type_of_contract = i.contents[1].string

                elif i.contents[0].string == "Очереди":
                    queues = i.contents[1].string

                elif i.contents[0].string == "Число корпусов":
                    number_of_buildings = i.contents[1].string

                elif i.contents[0].string == "Этажность":
                    storeys = i.contents[1].string

                elif i.contents[0].string == "Число квартир":
                    number_of_apartments = i.contents[1].string

                elif i.contents[0].string == "Высота потолков":
                    ceiling_height = i.contents[1].string

                elif i.contents[0].string == "Тип дома":
                    house_type = i.contents[1].string

                elif i.contents[0].string == "Отделка":
                    finish = i.contents[1].string

                elif i.contents[0].string == "Машиноместа в паркинге":
                    parking = i.contents[1].string

            advantages = ""

            try:
                Info3 = soup.find('div', {'class': 'CardFeatures__extra'}).contents[1]

                for i in Info3:
                    advantages += i.contents[1].contents[0].string + ", "
            except:
                None


            advantages += soup.find('div', {'class': 'SiteCardDescription__text'}).string


            images = ""
            Images = soup.find('div', {'class': 'GalleryThumbsSlider'})
            for image in Images:
                images+="https:" + str(image.contents[0]["src"]).replace("minicard","large")+","

            images = images[:-1]


            print("URL: " + url)
            print("Новый URL: " + url2)
            print("Название: "+name)
            print("Адрес: "+address)
            print("Метро: "+metro)
            print("Класс жилья: " + class_)
            print("Квартиры от: " + area_from)
            print("Квартиры до: " + area_to)
            print("Цена от: " + price_from)
            print("Цена до: " + price_to)
            print("Срок сдачи: " + deadline)
            print("Тип договора: " + type_of_contract)
            print("Очереди: " + queues)
            print("Число корпусов: " + number_of_buildings)
            print("Этажность: " + storeys)
            print("Число квартир: " + number_of_apartments)
            print("Высота потолков: " + ceiling_height)
            print("Тип дома: " + house_type)
            print("Отделка: " + finish)
            print("Машиноместа в паркинге: " + parking)
            print("Дополнительное описание: " + advantages)
            print("Изображения: " + images)

            print(str(id_now)+"/"+str(len(ids)))
            print("")


            #Запись в excel
            ws.cell(id_now+1, 1).value = url
            ws.cell(id_now+1, 2).value = url2
            ws.cell(id_now+1, 3).value = name
            ws.cell(id_now+1, 4).value = address
            ws.cell(id_now+1, 5).value = metro
            ws.cell(id_now+1, 6).value = class_
            ws.cell(id_now+1, 7).value = area_from
            ws.cell(id_now+1, 8).value = area_to
            ws.cell(id_now+1, 9).value = price_from
            ws.cell(id_now+1, 10).value = price_to
            ws.cell(id_now+1,  11).value = deadline
            ws.cell(id_now+1,  12).value = type_of_contract
            ws.cell(id_now+1,  13).value = queues
            ws.cell(id_now+1,  14).value = number_of_buildings
            ws.cell(id_now+1,  15).value = storeys
            ws.cell(id_now+1,  16).value = number_of_apartments
            ws.cell(id_now+1,  17).value = ceiling_height
            ws.cell(id_now+1,  18).value =  house_type
            ws.cell(id_now+1,  19).value = finish
            ws.cell(id_now+1, 20).value = parking
            ws.cell(id_now+1, 21).value = advantages
            ws.cell(id_now+1, 22).value = images


            #Запись в csv



            id_now += 1
            wb.save("Итог.xlsx")

            time.sleep(time_sleep)
        except:
            None




    for row in ws.rows:
        for el in row:

            csv_.write(str(el.value))
            csv_.write(";")
        csv_.write("\n")


    wb.close()
    csv_.close()


    print()
    print("Парсинг завершен\nНажмите Enter для выхода")
    input()

else:
    print("Файл <<ID.txt>> не найден")
    print("Нажмите Enter для выхода")
    input()