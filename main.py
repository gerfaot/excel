import xlsxwriter
import requests
from bs4 import BeautifulSoup
import pandas as pd
excel_data_df = pd.read_excel('ГРАФИК.xlsx', sheet_name='Мультиконтейнеры')
container = excel_data_df['№ контейнера'].tolist()
i = 1
workbook = xlsxwriter.Workbook('next.xlsx')
worksheet = workbook.add_worksheet()
exep = 'problem'
for item in container:
    try:
        response = requests.get('http://www.cma-cgm.com/ebusiness/tracking/search?SearchBy=Container&Reference=' + item)
        soup = BeautifulSoup(response.text, "lxml")
        worksheet.write(i, 4, item)
        worksheet.write(i, 5, soup.findAll("td", class_="is-header js-openrow")[-1].text.strip())
        worksheet.write(i, 6, soup.findAll("td", class_="is-headerdata js-openrow")[-1].text.strip())
        worksheet.write(i, 7, soup.findAll("td", attrs={'data-label': 'Vessel'})[-1].text.strip())
        print(i)
        i += 1
        print(soup.findAll("td", class_="is-header js-openrow")[-1].text.strip())
        print(soup.findAll("td", class_="is-headerdata js-openrow")[-1].text.strip())
        print(soup.findAll("td", attrs={'data-label': 'Vessel'})[-1].text.strip())
    except Exception as ex:
        print('exeption', ex)
        worksheet.write(i, 4, item)
        worksheet.write(i, 5, exep)
        print(i)
        i += 1
        continue
workbook.close()






#wb = openpyxl.load_workbook('Образец.xlsx') #Открываем тестовый Excel файл
#wb.create_sheet('Sheet1') #Создаем лист с названием "Sheet1"
#worksheet = wb['Sheet1'] #Делаем его активным
#worksheet['B4']='We are writing to B4' #В указанную ячейку на активном листе пишем все, что в кавычках
#wb.save('testdel.xlsx') #Сохраняем измененный файл

#workbook = xlsxwriter.Workbook('hello.xlsx')
#worksheet = workbook.add_worksheet()
#worksheet.write('K2', soup.findAll("td", class_="is-header js-openrow")[-1].text)
#worksheet.write(4, 3, 'Hello world')
#workbook.close()
#print(item)
#print(response.text) # вывод содержимого страницы