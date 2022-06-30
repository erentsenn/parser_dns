import pandas
from selenium.webdriver.common.by import By
from selenium import webdriver
from pathlib import Path
from selenium.webdriver import Firefox, FirefoxProfile
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
import time

BASEDIR = Path(__file__).parent


class Profile(FirefoxProfile):
    accept_untrusted_certs = True
    assume_untrusted_cert_issuer = False

    def __init__(self, profile_directory=None):
        super().__init__(profile_directory)
        # отключаем WebRTC
        self.set_preference("media.peerconnection.enabled", False)
        self.set_preference("dom.webnotifications.enabled", False)
        self.set_preference("dom.webdriver.enabled", False)
        self.set_preference("browser.aboutConfig.showWarning", False)


class Client(Firefox):
    def __init__(self):
        profile = Profile()
        super().__init__(firefox_profile=profile,
                         executable_path='geckodriver.exe')


dataframe = pandas.DataFrame({"category": [],
                              "man": [],
                              "name": [],
                              "mod": [],
                              "color": [],
                              "color_ru": []
                              })
available_cat = ['Ноутбук',
                 'Ультрабук']
available_man = ['4Good',
                 'Acer',
                 'Alienware',
                 'AORUS',
                 'Apple',
                 'ASUS',
                 'BB-mobile',
                 'Chuwi',
                 'Dell',
                 'DEXP',
                 'Digma',
                 'Dream Machines',
                 'Echips',
                 'Fujitsu',
                 'GIGABYTE',
                 'Haier',
                 'HONOR',
                 'HP',
                 'HUAWEI',
                 'Infinix',
                 'Irbis',
                 'Krez',
                 'Lenovo',
                 'LG',
                 'Mitac',
                 'Getac',
                 'MSI',
                 'MX',
                 'Packard Bell',
                 'Panasonic',
                 'Prestigio',
                 'Samsung',
                 'Sony',
                 'Toshiba',
                 'Xiaomi',
                 'ZET GAMING',
                 'ДНС (59)']
available_mod_ram = ['RAM 2 ГБ',
                     'RAM 4 ГБ',
                     'RAM 6 ГБ',
                     'RAM 8 ГБ',
                     'RAM 12 ГБ',
                     'RAM 16 ГБ',
                     'RAM 24 ГБ',
                     'RAM 32 ГБ',
                     'RAM 64 ГБ']
available_mod_hdd = ['HDD 120 ГБ',
                     'HDD 128 ГБ',
                     'HDD 240 ГБ'
                     'HDD 256 ГБ',
                     'HDD 500 ГБ',
                     'HDD 512 ГБ',
                     'HDD 1000 ГБ',
                     'HDD 1024 ГБ',
                     'HDD 1536 ГБ',
                     'HDD 2000 ГБ',
                     'HDD 4000 ГБ',
                     'HDD 64 ГБ']
available_mod_ssd = ['SSD 120 ГБ',
                     'SSD 128 ГБ',
                     'SSD 240 ГБ'
                     'SSD 256 ГБ',
                     'SSD 500 ГБ',
                     'SSD 512 ГБ',
                     'SSD 1000 ГБ',
                     'SSD 1024 ГБ',
                     'SSD 1536 ГБ',
                     'SSD 2000 ГБ',
                     'SSD 4000 ГБ',
                     'SSD 64 ГБ']
available_color_dict = {'бежевый': "beige",
                        'белый': "white",
                        'бронзовый': "bronze",
                        'голубой': "blue",
                        'зеленый': "green",
                        'золотистый': "golden",
                        'коричневый': "brown",
                        'красный': "red",
                        'перламутровый': "pearl",
                        'розовый': "pink",
                        'серебристый': "silver",
                        'серый': "gray",
                        'синий': "blue",
                        'сиреневый': "lilac",
                        'фиолетовый': "purple",
                        'черный': "black"}
driver = Client()
driver.get('https://'
           'www.dns-shop.ru/'
           'catalog/'
           '17a892f816404e77/'
           'noutbuki/'
           '?stock=now-today-tomorrow-later-out_of_stock&'
           'mode=simple')

for page in range(500):
    for i in range(1, 10):
        try:
            for j in range(1, 20):
                try:
                    text = driver.find_element(by=By.XPATH,
                                               value=f'/html/body/div[1]/div/div[2]/div[2]/div[3]/div/div[{i}]/div[{j}]/a/span').text
                    category_of_laptop = ''
                    man_of_laptop = ''
                    name_of_laptop = ''
                    mod_of_laptop = ''
                    color_of_laptop = ''
                    color_ru_of_laptop = ''
                    mod_string = []

                    for ctg in available_cat:
                        if ctg in text:
                            category_of_laptop = ctg
                    for man in available_man:
                        if man in text:
                            man_of_laptop = man
                    for ram in available_mod_ram:
                        if ram in text:
                            mod_string.append(ram)
                    for hdd in available_mod_hdd:
                        if hdd in text:
                            mod_string.append(hdd)
                    for ssd in available_mod_ssd:
                        if ssd in text:
                            mod_string.append(ssd)
                    for color_ru, color_eng in available_color_dict.items():
                        if color_ru in text.lower() or color_eng in text.lower():
                            color_of_laptop = color_eng
                            color_ru_of_laptop = color_ru
                    try:
                        lst = text.split()
                        start_index = lst.index(category_of_laptop)
                        end_index = lst.index(color_ru_of_laptop)
                        name_of_laptop = ' '.join(lst[start_index + 1: end_index])
                    except Exception as e:
                        print(e)
                        print(color_ru_of_laptop, category_of_laptop)
                        name_of_laptop = 'indefinite'

                    dict_to_df = {"category": category_of_laptop,
                                  "man": man_of_laptop,
                                  "name": name_of_laptop,
                                  "mod": ' '.join(mod_string),
                                  "color": color_of_laptop,
                                  "color_ru": color_ru_of_laptop
                                  }
                    print(dict_to_df)
                    dataframe = dataframe.append(dict_to_df, ignore_index=True)
                except Exception as e:
                    continue
        except Exception as e:
            continue
    try:
        print(dataframe)
        driver.get(f'https://www.dns-shop.ru/catalog/17a892f816404e77/noutbuki/?stock=now-today-tomorrow-later-out_of_stock&p={page}&mode=simple')
        time.sleep(5)
    except Exception as e:
        pass
writer = pd.ExcelWriter('output.xlsx')
dataframe.to_excel(writer)
writer.save()