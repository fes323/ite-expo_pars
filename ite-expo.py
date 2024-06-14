import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def get_driver():
    proxy = 'http://10.250.0.1:3128'
    options = Options()
    options.add_argument(f'--proxy-server={proxy}')
    options.add_argument('--window-size=1920,1080')
    driver = webdriver.Chrome(options=options)
    return driver

def parsing_list_page(driver, url):
    driver.get(url)

    mainCategories = {
        48243: '03. Упаковочное оборудование',
        48297: '04. Готовая упаковка и одноразовая посуда',
        48372: '05. Этикетка',
        48376: '06. Транспортная упаковка'
    }

    additionalCategories = {
        3425: 'Замороженные продукты, полуфабрикаты и готовая кулинария',
        3422: 'Бакалея, кондитерские или хлебобулочные изделия',
    }

    mainCategoriesFilterId = 'p_lt_zoneContainer_pageplaceholder_p_lt_zoneForm_Filter_filterControl_ddlMainCategories'
    additionalCategoriesFilterId = 'p_lt_zoneContainer_pageplaceholder_p_lt_zoneForm_Filter_filterControl_ddlAdditionalCategories'
    searchButtonId = 'p_lt_zoneContainer_pageplaceholder_p_lt_zoneForm_Filter_filterControl_fltrbtn'

    detail_urls_list = []

    for mCategory in mainCategories:

        driver.get(url)

        # Выбираем основной рубрикатор
        driver.find_element(By.ID, mainCategoriesFilterId).click()
        driver.find_element(By.ID, mainCategoriesFilterId).send_keys(mainCategories[mCategory])

        # Запускаем поиск и ждем рандомное кол-во секунд от 0 до 4
        driver.find_element(By.ID, searchButtonId).click()
        time.sleep(random.randint(0, 4))

        url_category = mainCategories.get(mCategory)
        break_while = False
        while True:

            if break_while == True:
                break
            # Парсим url адреса страниц с детальной информацией о компании
            time.sleep(random.randint(0, 3))
            raw_urls = driver.find_elements(By.CLASS_NAME, 'popUp')
            for raw_url in raw_urls:
                data_list = []
                try:
                    country = raw_url.find_element(By.CLASS_NAME, 'country').text
                except:
                    country = 'None'
                data_url = raw_url.get_attribute('href')
                data_list.append(data_url)
                data_list.append(country)
                if {url_category: data_list} in detail_urls_list:
                    break_while = True
                    break
                detail_urls_list.append({url_category:data_list})

            next_btn = driver.find_element(By.XPATH, "//a[contains(text(), '>')]")
            next_btn.click()

    for aCategory in additionalCategories:

        driver.get(url)

        # Выбираем основной рубрикатор
        driver.find_element(By.ID, additionalCategoriesFilterId).click()
        driver.find_element(By.ID, additionalCategoriesFilterId).send_keys(additionalCategories[aCategory])

        # Запускаем поиск и ждем рандомное кол-во секунд от 0 до 4
        driver.find_element(By.ID, searchButtonId).click()
        time.sleep(random.randint(0, 5))

        url_category = additionalCategories.get(aCategory)
        break_while = False
        while True:

            if break_while == True:
                break
            # Парсим url адреса страниц с детальной информацией о компании
            time.sleep(random.randint(0, 3))
            raw_urls = driver.find_elements(By.CLASS_NAME, 'popUp')
            for raw_url in raw_urls:
                data_list = []
                try:
                    country = raw_url.find_element(By.CLASS_NAME, 'country').text
                except:
                    country = 'None'
                data_url = raw_url.get_attribute('href')
                data_list.append(data_url)
                data_list.append(country)
                if {url_category: data_list} in detail_urls_list:
                    break_while = True
                    break
                detail_urls_list.append({url_category:data_list})

            next_btn = driver.find_element(By.XPATH, "//a[contains(text(), '>')]")
            time.sleep(random.randint(0, 3))
            next_btn.click()

    return detail_urls_list

def parsing_detail_page(detail_urls_list, driver):
    data_to_write = []
    for detail_dict_url in detail_urls_list:
        for key, value in detail_dict_url.items():
            category = key
            detail_url = value[0]
            country = value[1]

        driver.get(str(detail_url))

        try:
            company_name = driver.find_element(By.XPATH, '/html/body/form/div[5]/div/div[1]/h2').text
        except:
            company_name = ''
        try:
            company_adres = driver.find_element(By.XPATH, '/html/body/form/div[5]/div/div[3]/p').text
        except:
            company_adres = ''
        try:
            company_phone = driver.find_element(By.XPATH, '/html/body/form/div[5]/div/div[5]/div/p').text
        except:
            company_phone = ''
        try:
            company_website = driver.find_element(By.XPATH, "//a[contains(@href, '//')]")
            company_website = company_website.get_attribute('href')
            print(company_website)
        except:
            company_website = ''
        try:
            company_email = driver.find_element(By.XPATH, "//a[contains(@href, 'mailto:')]")
            company_email = company_email.get_attribute('href')
            print(company_email)
        except:
            company_email = ''

        data_to_write.append(
            {
                'Категория': category,
                'Название': company_name,
                'Страна': country,
                'Адрес': company_adres,
                'Телефон': company_phone,
                'Сайт': company_website,
                'Email': company_email
            }
        )
    return data_to_write

def write_to_xlsx(data_to_write):
    df = pd.DataFrame(data_to_write)
    with pd.ExcelWriter('RosUpack.xlsx') as writer:
        df.to_excel(writer)
    return 0


if __name__ == '__main__':
    driver = get_driver()
    detail_urls_list = parsing_list_page(driver, 'https://catalogue.ite-expo.ru/ru-RU/exhibitorlist.aspx?project_id=521')
    data_to_write = parsing_detail_page(detail_urls_list, driver)
    write_to_xlsx(data_to_write)