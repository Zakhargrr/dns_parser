from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import pandas as pd
from seleniumwire import webdriver
import time
from selenium.webdriver.chrome.service import Service

cookies = '' # введите валидные куки
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Cookie': cookies,
}

options = Options()
options.page_load_strategy = 'normal'


def append_to_excel(filepath, df, sheet_name):
    try:
        with pd.ExcelWriter(filepath, mode="a") as f:
            df.to_excel(f, sheet_name=sheet_name, index=False)
    except FileNotFoundError:
        with pd.ExcelWriter(filepath, mode="w") as f:
            df.to_excel(f, sheet_name=sheet_name, index=False)


def interceptor(request):
    del request.headers['User-Agent']
    request.headers['User-Agent'] = headers['User-Agent']

    del request.headers['Accept']
    request.headers['Accept'] = headers['Accept']

    del request.headers['Cookie']
    request.headers['Cookie'] = headers['Cookie']


def print_data(product_dict):
    images_string = product_dict['Ссылки на все изображения']
    data_string = f"\n\n{product_dict['Наименование']}\nЦена: {product_dict['Цена']} руб.\n{product_dict['Продается/Не продается']}\n{product_dict['Ссылка на товар']}\nГлавное изображение: {product_dict['Ссылка на главное изображение']}\nВсе изображения\n{images_string}\n{product_dict['Характеристики']}\n{product_dict['Описание']}"
    print(data_string)


def parse_category(category: str):
    products = []
    page = 1
    counter = 1

    driver = webdriver.Chrome(service=Service('D:\\SkyPro_2\\dns_parser\\chromedriver\\chromedriver.exe'),
                              options=options)
    driver.request_interceptor = interceptor

    products_amount = None
    # если нужно ограничить число записей в таблице
    # products_amount = 30

    if products_amount is None:
        category_url = f'https://www.dns-shop.ru/search/?q={category}&p={page}&order=popular&stock=all/'
        driver.get(url=category_url)
        for request in driver.requests:
            if 'https://www.dns-shop.ru/search/' in request.url:
                if request.response.status_code != 200:
                    cookies = input('Куки перестали быть валидными\n\nВведите новые куки (https://www.dns-shop.ru/): ')
                    headers['Cookie'] = cookies
                    driver.request_interceptor = interceptor
                    driver.get(url=category_url)

        category_soup = BeautifulSoup(driver.page_source, 'html.parser')
        products_amount_str = category_soup.find('div', class_="tabs-top-filters__slider-wrapper tns-item tns-slide-active").find('div',                                                                                                      class_="tabs-top-filters__count").contents[0]
        products_amount = int(str(products_amount_str))

    while counter <= products_amount:
        category_url = f'https://www.dns-shop.ru/search/?q={category}&p={page}&order=popular&stock=all/'
        driver.get(url=category_url)
        for request in driver.requests:
            if 'https://www.dns-shop.ru/search/' in request.url:
                if request.response.status_code != 200:
                    cookies = input('Куки перестали быть валидными\n\nВведите новые куки (https://www.dns-shop.ru/): ')
                    headers['Cookie'] = cookies
                    driver.request_interceptor = interceptor
                    driver.get(url=category_url)
        category_soup = BeautifulSoup(driver.page_source, 'html.parser')

        links = category_soup.find_all('a', class_="catalog-product__name ui-link ui-link_black")
        for link in links:
            product_url = 'https://www.dns-shop.ru' + link['href']
            driver.get(url=product_url)
            for request in driver.requests:
                if 'https://www.dns-shop.ru/product/' in request.url:
                    if request.response.status_code != 200:
                        cookies = input('Куки перестали быть валидными\n\nВведите новые куки(https://www.dns-shop.ru/): ')
                        headers['Cookie'] = cookies
                        driver.request_interceptor = interceptor
                        driver.get(url=product_url)

            driver.find_element(By.CLASS_NAME, 'product-characteristics__expand').click()
            time.sleep(3)
            product_soup = BeautifulSoup(driver.page_source, 'html.parser')
            title = str(product_soup.find('div', class_="product-card-top__name").contents[0])
            try:
                price_str = product_soup.find('div', class_="product-buy__price product-buy__price_active").contents[0][:-2]
            except AttributeError:
                price_str = product_soup.find('div', class_="product-buy__price").contents[0][:-2]
            price = int(str(price_str).replace(' ', ''))
            try:
                shops_info = str(product_soup.find('span', class_="available").contents[0])
            except AttributeError:
                availability = 'Не продается'
            else:
                if shops_info == "В наличии: ":
                    availability = 'Продается'
                else:
                    availability = 'Не продается'

            main_image_tag = product_soup.find('img', class_="product-images-slider__img")
            try:
                main_image = str(main_image_tag['src'])
            except KeyError:
                main_image = str(main_image_tag['data-src'])

            images = product_soup.find_all('img', class_="product-images-slider__img loaded tns-complete")
            for i in range(len(images)):
                try:
                    images[i] = images[i]['src']
                except KeyError:
                    images[i] = str(images[i]['data-src'])
            groups = product_soup.find_all('div', class_="product-characteristics__group")
            characteristics = ""
            for group in groups:
                group_title = group.find('div', class_="product-characteristics__group-title")
                characteristics += str(group_title.contents[0]) + ': '
                group_spec_titles = group.find_all('div', class_="product-characteristics__spec-title")
                group_value_titles = group.find_all('div', class_="product-characteristics__spec-value")
                length = len(group_spec_titles)
                for i in range(len(group_spec_titles)):
                    if len(group_value_titles[i]) > 1:
                        try:
                            group_value_titles[i].contents[1].contents
                        except AttributeError:
                            characteristics += str(group_spec_titles[i].contents[0]).replace('\t', '') + ' - ' + str(
                                group_value_titles[i].contents[1]).replace('\t', '').replace('\n', '')
                        else:
                            characteristics += str(group_spec_titles[i].contents[0]).replace('\t', '') + ' - ' + str(
                                group_value_titles[i].contents[1].contents[0]).replace('\t', '').replace('\n', '')
                    else:
                        try:
                            group_value_titles[i].contents[0].contents
                        except AttributeError:
                            characteristics += str(group_spec_titles[i].contents[0]).replace('\t', '') + ' - ' + str(
                                group_value_titles[i].contents[0]).replace('\t', '').replace('\n', '')
                        else:
                            characteristics += str(group_spec_titles[i].contents[0]).replace('\t', '') + ' - ' + str(
                                group_value_titles[i].contents[0].contents[0]).replace('\t', '').replace('\n', '')
                    if i < length - 1:
                        characteristics += ';'
                    else:
                        characteristics += '.\n\n'

            description_tag = product_soup.find('div', class_="product-card-description-text")
            description = '\n'.join(list(map(str, description_tag.p.contents))).replace('<br>', '').replace('<br/>', '')
            product_dict = {
                'Категория': category,
                'Наименование': title,
                'Цена': price,
                'Продается/Не продается': availability,
                'Ссылка на товар': product_url,
                'Ссылка на главное изображение': main_image,
                'Ссылки на все изображения': ' \n'.join(images),
                'Характеристики': characteristics,
                'Описание': description,
            }
            products.append(product_dict)
            print_data(product_dict)
            counter += 1
            if counter > products_amount:
                break
        page += 1

    df = pd.DataFrame(data=products)
    # если записать в отдельный excel-файл
    # df.to_excel(f"xlsx_files/{category}.xlsx", index=False)

    # если добавить добавить все данные в один excel-файл
    append_to_excel(f"xlsx_files/parsed_data.xlsx", df, f"{category}")
    driver.close()
    driver.quit()


category = input("Введите категорию товара для парсинга: ")
parse_category(category)
