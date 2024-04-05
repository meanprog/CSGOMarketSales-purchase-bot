import datetime
import requests
import openpyxl



apiKey = "key" #апи ключ
dateStart = "31-12-2023" #начальная дата
dateEnd = "03-04-2024" #конечная дата
date_time_Start = datetime.datetime.strptime(dateStart, '%d-%m-%Y')
date_time_End = datetime.datetime.strptime(dateEnd, '%d-%m-%Y')

buy = [] #куплено/продано
params = {"key" : apiKey} # параметры



try:
    while(date_time_Start <= date_time_End):
        print(date_time_Start)
        #увеличиваем день
        date_time_Start += datetime.timedelta(days=1)
        #меняем ключ в параметрах к запросу
        params['date'] = str(date_time_Start.date().strftime('%d-%m-%Y'))
        #делаем запрос
        response = requests.get("https://market.csgo.com/api/v2/history?", params=params)
        print(response.status_code)
        while(response.status_code != 200):
            response = requests.get("https://market.csgo.com/api/v2/history?", params=params)
        #получаем json

        dataMarket = response.json()

        #добавляем в массив куплено элементы из json
        for item in dataMarket['data']:
            if (str(item['stage']) == '2'):
                if(str(item['event']) == 'sell'):
                    buy.append([date_time_Start, item['market_hash_name'], item['received'], item['event']])
                elif(str(item['event']) == 'buy'):
                    buy.append([date_time_Start, item['market_hash_name'], item['paid'],  item['event']])


    #проходим по массиву и ищем одинаковые элементы, но с разными ивентами
    for i in buy:
        for j in buy:
            if i[1] == j[1] and i[3] == "buy" and j[3] == "sell" and 0 <= 3 < len(i) and (j[0] >= i[0] + datetime.timedelta(days=8)):
                i.append(j[0])
                i.append(j[3])
                i.append(j[2])
                i.append(int(i[6])-int(i[2]))
                buy.remove(j)
                break

    #print(buy)


    #массив хранит информацию и о нереализованных товарах, чтобы получить полный отчёт нужно отредактировать if в условиях вывода
    #записываем в эксель
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = 'Дата покупки'
    sheet['B1'] = 'Дата продажи'
    sheet['C1'] = 'Нэйм'
    sheet['D1'] = 'Купил'
    sheet['E1'] = 'Продал'
    sheet['F1'] = 'Бизнес'
    row = 2
    sum = 0
    for item in buy:
        if 0 <= 5 < len(item):
            sheet[row][0].value = str(item[0].date().strftime('%d.%m.%Y'))
            sheet[row][1].value = str(item[4].date().strftime('%d.%m.%Y'))
            sheet[row][2].value = str(item[1])
            if (item[3] == 'sell'):
                sheet[row][4].value = int(item[2]) / 100
            elif ((item[3] == 'buy')):
                sheet[row][3].value = int(item[2]) / 100
            sheet[row][4].value = int(item[6])/100
            sheet[row][5].value = int(item[7])/100
            sum += int(item[7])/100
            row += 1
        # elif 0 <= 3 < len(item):
        #     if ((item[3] == 'buy')):
        #         sheet[row][0].value = str(item[0].date().strftime('%d.%m.%Y'))
        #         sheet[row][2].value = str(item[1])
        #         sheet[row][3].value = int(item[2]) / 100
        #         row += 1
    sheet[row][0].value = "Итог: " + str(sum)
    book.save("BuySell.xlsx")
    book.close()


except Exception as ex:
    print(ex)


