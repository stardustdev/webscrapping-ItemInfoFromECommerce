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
        self.main_url = "https://www.multifoods.com.br/produtos"
        self.browser = webdriver.Chrome()
        self.browser.get(self.main_url)
        self.browser.maximize_window()

        # connect to google sheet
        scope = ['https://www.googleapis.com/auth/drive', 'https://spreadsheets.google.com/feeds']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('SheetAndPython-8d3b50000138.json', scope)
        self.client = gspread.authorize(self.credentials)
        self.sheet = self.client.open('MultifoodsBOT').sheet1

        # parameters
        self.products_url = []
        self.products_num = 0

        self.page_num = 0
        self.urls = []
        self.work_finished = False

    def get_total_num_of_products(self):

        # stage for login
        time.sleep(5)

        self.browser.find_element_by_partial_link_text("Faça o seu login").click()
        time.sleep(2)
        while True:
            try:
                self.browser.find_element_by_class_name("simple-field").send_keys("floki.tech.contato@gmail.com")
                self.browser.find_element_by_name("senha").send_keys("flokitech2019")
                self.browser.find_element_by_xpath("/html/body/div/div[1]/div[2]/div/div[1]/div/div[1]/div/form/div[2]/input").click()

                time.sleep(4)
            except:
                continue
            break

        while True:
            try:
                # stage for getting total product counter and page counter
                total_number = self.browser.find_element_by_class_name("description").text
                self.products_num = int(re.findall(r'[-+]?\d*\.\d+|\d+', total_number)[2])
                self.page_num = int((self.products_num + 30) / 60)
                print(
                    "product number : " + str(self.products_num) + " and " + "page number : " + str(self.page_num))
            except:
                continue
            break
    def get_url_of_each_product(self):
        for i in range(0, self.page_num):
            while True:
                try:
                    elements = self.browser.find_elements_by_class_name("col-md-3.col-sm-4.shop-grid-item")
                    for ele in elements:
                        # get url of each product
                        category = ele.find_element_by_class_name("tag.categoria")
                        url = category.get_attribute("href")
                        category = category.text
                        self.products_url.append({
                            'url':url,
                            'category':category
                        })
                except:
                    continue
                break

            # click pagination button to go to next page
            if (i != self.page_num - 1):
                while True:
                    try:
                        self.browser.find_element_by_css_selector(".pages-box > a[data-index*='" + str(i + 2) + "']").click()
                        time.sleep(3)
                    except:
                        continue
                    break

        # output into google sheet(url and category)
        while True:
            try:
                cell_list = self.sheet.range(2, 1, 2+self.products_num, 1)
                for j, val in enumerate(self.products_url):  # gives us a tuple of an index and value
                    cell_list[j].value = val['url']  # use the index on cell_list and the val from cell_values

                self.sheet.update_cells(cell_list)

                cell_list = self.sheet.range(2, 2, 2+self.products_num, 2)
                for j, val in enumerate(self.products_url):  # gives us a tuple of an index and value
                    cell_list[j].value = val['category']  # use the index on cell_list and the val from cell_values

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
        print(self.products_url)

    def get_information_from_each_product(self):
        print("here is extract_info_from_each_product")

        url_index = 0
        self.sheet.update_cell(2, 12, datetime.datetime.today().strftime("%d-%b-%Y"))

        while url_index < self.products_num:
            while True:
                try:
                    url = self.products_url[url_index]['url']

                    if (url == None) or (url == ""):  # If url doesn't exist, iteration is finished
                        self.work_finished = True
                        break

                    print(url)

                    self.browser.get(url)
                    self.browser.maximize_window()
                    time.sleep(2.5)

                    product_body = self.browser.find_element_by_class_name('product-detail-box')

                    try:
                        name = product_body.find_element_by_class_name("product-title").text
                    except:
                        name = ''

                    try:
                        code = self.browser.find_element_by_xpath("//*[contains(text(), 'Código: ')]").text.replace("Código: ", "")
                    except:
                        code = ""

                    try:
                        barcode = self.browser.find_element_by_xpath("//*[contains(text(), 'Código de Barras: ')]").text.replace("Código de Barras: ", "")
                    except:
                        barcode = ""

                    try:
                        marca = self.browser.find_element_by_xpath("//*[contains(text(), 'Marca: ')]").text.replace("Marca: ", "")
                    except:
                        marca = ""

                    try:
                        unit = self.browser.find_element_by_xpath("//*[contains(text(), 'Unidade: ')]").text.replace("Unidade: ", "")
                    except:
                        unit = ""

                    try:
                        amount = self.browser.find_element_by_xpath("//*[contains(text(), 'Peso Aproximado: ')]").text.replace("Peso Aproximado: ", "")
                    except:
                        amount = ""

                    try:
                        priceperunit = self.browser.find_element_by_xpath("//*[contains(text(), 'Preço por')]").text
                        priceperunit = re.findall(r'[-+]?\d*,\d+|\d+', priceperunit)[0]
                    except:
                        priceperunit = ""

                    try:
                        price = product_body.find_element_by_class_name("current").text.replace('R$ ', '')
                    except:
                        price = ""

                    # output into google sheet
                    while True:
                        try:
                            print('step 1')
                            start_time = datetime.datetime.now()
                            cell_list = self.sheet.range(url_index+2, 3, url_index+2, 10)
                            end_time = datetime.datetime.now()
                            if (end_time - start_time).total_seconds() < 1:
                                time.sleep(1.01 - (end_time - start_time).total_seconds())
                            print('step 2')
                            cell_values = [name, price, priceperunit, unit, amount, marca, code, barcode]
                            print(cell_values)
                            print('step 3')
                            for j, val in enumerate(cell_values):  # gives us a tuple of an index and value
                                cell_list[j].value = val  # use the index on cell_list and the val from cell_values

                            print('step 4')
                            start_time = datetime.datetime.now()
                            self.sheet.update_cells(cell_list)
                            end_time = datetime.datetime.now()
                            if (end_time - start_time).total_seconds() < 1:
                                time.sleep(1.01 - (end_time - start_time).total_seconds())

                            print(str(self.products_num) + "/" + str(url_index + 1))
                            url_index += 1
                        except:
                            print("normal except")
                            continue
                        break

                except:
                    continue
                break


if __name__ == "__main__":

    ExtractProduct = ExtractProduct()
    ExtractProduct.get_total_num_of_products()
    ExtractProduct.get_url_of_each_product()
    ExtractProduct.get_information_from_each_product()
    print("Task completed")