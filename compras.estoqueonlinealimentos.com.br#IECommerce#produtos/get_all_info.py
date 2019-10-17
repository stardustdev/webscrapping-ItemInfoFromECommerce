from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import load_workbook
import gspread
import datetime
import time
import urllib
from bs4 import BeautifulSoup
from urllib.request import Request
import re

class ExtractProduct():

    def __init__(self):
        # set chrome webdriver
        self.main_url = "https://compras.estoqueonlinealimentos.com.br/IECommerce/produtos"
        self.browser = webdriver.Chrome()
        self.browser.get(self.main_url)
        self.browser.maximize_window()
        self.browser.refresh()
        time.sleep(2)

        # connect to google sheet
        scope = ['https://www.googleapis.com/auth/drive', 'https://spreadsheets.google.com/feeds']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('SheetAndPython-8d3b50000138.json', scope)
        self.client = gspread.authorize(self.credentials)
        self.sheet = self.client.open('EstoqueOnlineBOT').sheet1

        # parameters
        self.categories = ['ATOMATADOS', 'BEBIDAS', 'BOMBONIERE',
                         'CEREAIS', 'CHAS', 'CONDIMENTOS',
                         'CONSERVAS', 'DIET', 'DOCES',
                         'ENLATADOS', 'ESSENCIAS', 'FARINACEOS',
                         'FRUTAS SECAS / CALDA', 'GRAOS SECOS', 'LEITE E DERIVADOS',
                         'MAIONESES', 'MARGARINAS', 'MASSAS',
                         'MATINAIS', 'MOLHOS', 'OLEOS E AZEITES',
                         'PANIFICADORA E CONF', 'SACHET', 'SEMENTES',
                         'SNACKS', 'SOBREMESAS']

        self.products_url = []
        self.products_num = 0

        self.page_num = 0
        self.urls = []
        self.work_finished = False

    def get_total_num_of_products(self):
        # stage for getting total product counter and page counter
        total_number = self.browser.find_element_by_class_name("description").text
        self.products_num = int(re.findall(r'[-+]?\d*\.\d+|\d+', total_number)[2])
        self.page_num = int((self.products_num+30)/60)
        print("product number : " + str(self.products_num) + " and " + "page number : " + str(self.page_num))

    def get_information_from_each_product(self):
        self.sheet.update_cell(2, 9, datetime.datetime.today().strftime("%d-%b-%Y"))
        for category in self.categories:
            self.browser.find_element_by_link_text(category).click()
            time.sleep(1)

            scroll_ele = self.browser.find_element_by_xpath(
                "//div[@class='ui-datascroller-content ui-widget-content ui-corner-bottom']")
            # # time.sleep(1)
            no_of_pagedowns = 40

            while no_of_pagedowns:
                self.browser.execute_script("arguments[0].scrollBy(0, 600);", scroll_ele)
                time.sleep(0.2)
                no_of_pagedowns -= 1

            all_products = self.browser.find_elements_by_class_name("w3-col.l2.m6.s12.card")
            print(len(all_products))
            for i in range(len(all_products)):
                scroll_ele = self.browser.find_element_by_xpath("//div[@class='ui-datascroller-content ui-widget-content ui-corner-bottom']")
                # # time.sleep(1)
                no_of_pagedowns = (i//6)*10

                while no_of_pagedowns:
                    self.browser.execute_script("arguments[0].scrollBy(0, 600);", scroll_ele)
                    time.sleep(0.2)
                    no_of_pagedowns -= 1
                while True:
                    try:
                        product = self.browser.find_elements_by_class_name("w3-col.l2.m6.s12.card")[i]
                        product.find_element_by_tag_name('a').click()
                        time.sleep(1)
                        infos = self.browser.find_element_by_class_name("ui-tabs-panel.ui-widget-content.ui-corner-bottom").text.split("\n")
                        # output google sheet
                        print(infos)
                    except:
                        print("Here is a problem.")
                        continue
                    break


                while True:
                    try:
                        cell_list = self.sheet.range(self.products_num + 2, 2, self.products_num + 2, 7)
                        cell_values = [infos[0], category, infos[1].replace('Código do Produto: ', ''), infos[2].replace('EAN: ', ''), infos[3].replace('Embalagem de Vendas: ', ''), infos[4].replace('Valor Unitário: R$ ', '')]
                        for i, val in enumerate(cell_values):  # gives us a tuple of an index and value
                            cell_list[i].value = val  # use the index on cell_list and the val from cell_values

                        self.products_num += 1
                        self.sheet.update_cells(cell_list)
                    except:
                        if self.credentials.access_token_expired:
                            print("Here is except token expired")
                            self.client.login()  # refreshes the token
                            continue
                        else:
                            print("normal except")
                            continue
                    break
                self.browser.execute_script("window.history.go(-1)")
                time.sleep(1)

if __name__ == "__main__":

    ExtractProduct = ExtractProduct()
    ExtractProduct.get_information_from_each_product()
    print("Task completed")
