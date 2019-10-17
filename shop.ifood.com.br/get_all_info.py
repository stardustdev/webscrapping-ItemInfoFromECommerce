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
        self.category_url = [
            {
                'url': 'https://shop.ifood.com.br/mercearia',
                'category': 'mercearia'
            },
            {
                'url': 'https://shop.ifood.com.br/bebidas',
                'category': 'bebidas'
            },
            {
                'url': 'https://shop.ifood.com.br/queijos-e-frios',
                'category': 'queijos e frios'
            },
            {
                'url': 'https://shop.ifood.com.br/acougue-e-peixaria',
                'category': 'acougue e peixaria'
            },
            {
                'url': 'https://shop.ifood.com.br/embalagens',
                'category': 'embalagens'
            },
            {
                'url': 'https://shop.ifood.com.br/mais-categorias',
                'category': 'mais categorias'
            }
        ]
        self.num_page_down = [1000, 1200, 500, 500, 1000, 500]
        # connect to google sheet
        scope = ['https://www.googleapis.com/auth/drive', 'https://spreadsheets.google.com/feeds']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('SheetAndPython-8d3b50000138.json', scope)
        self.client = gspread.authorize(self.credentials)
        self.sheet = self.client.open('ProductFromShop').sheet1

        self.products_num = 0


    def goto_mainpage(self, url):
        print("here is goto_mainpage")
        self.browser = webdriver.Chrome()
        self.browser.get(url)

        self.browser.find_element_by_class_name("button.button--large.mrg-r-16").click()
        self.browser.find_element_by_id("cep-value").send_keys("05448-000")
        self.browser.find_element_by_id("cep-button").click()
        time.sleep(3)

        return True

    def show_all_products(self, no):
        print("here is show_all_products")
        elem_body = self.browser.find_element_by_tag_name("body")

        # down the scroll of main body of page programically
        no_of_pagedowns = self.num_page_down[no]
        while no_of_pagedowns:
            elem_body.send_keys(Keys.PAGE_DOWN)
            time.sleep(0.2)
            print(no_of_pagedowns)
            no_of_pagedowns -= 1

    def extract_each_product_url_and_save(self):
        print("here is extract_each_product_url_and_save")
        self.sheet.update_cell(2, 8, datetime.datetime.today().strftime("%d-%b-%Y"))

        for x in range(0, 6):
            try:
                if (self.goto_mainpage(self.category_url[x]['url'])):
                    self.show_all_products(x)

                # pick up information url of each product
                products = self.browser.find_elements_by_class_name("cardMainWrap.pad-16.pad-8-mob")

                for product in products:
                    url_each_pro = product.find_element_by_tag_name("a").get_attribute("href")
                    try:
                        name = product.find_element_by_class_name('nome').text
                        name = ' '.join(name.split())
                    except:
                        name = ''
                    try:
                        amount = ' '.join(product.find_element_by_class_name('quantidade').text.split())
                        amount = re.findall(r'[-+]?\d*\.\d+|\d+', amount)[0]
                    except:
                        amount = ''
                    try:
                        price = ' '.join(product.find_element_by_class_name('price').text.split()).replace('R$ ', '')
                    except:
                        price = ''
                    try:
                        priceperunit = ' '.join(product.find_element_by_class_name('unidades').text.split()).replace('R$ ', '')
                    except:
                        priceperunit = ''

                    # output google sheet
                    while True:
                        try:
                            cell_list = self.sheet.range(self.products_num + 2, 1, self.products_num + 2, 6)
                            cell_values = [url_each_pro, name, self.category_url[x]['category'], price, priceperunit, amount]
                            print(cell_values)
                            for i, val in enumerate(cell_values):  # gives us a tuple of an index and value
                                cell_list[i].value = val  # use the index on cell_list and the val from cell_values

                            self.sheet.update_cells(cell_list)
                            self.products_num += 1
                        except:
                            if self.credentials.access_token_expired:
                                print("Here is except token expired")
                                self.client.login()  # refreshes the token
                                continue
                            else:
                                print("normal except")
                                continue
                        break
            except:
                print("There is a issue from extracting each product url in the category : " + self.category_url[x]['category'])
                continue

        print(self.products_num)

    def extract_info_from_each_product(self):
        print("here is extract_info_from_each_product")

        url_index = 2

        while url_index < self.products_num:
        # while url_index > 0:
            while True:
                try:
                    print("Current Product No : " + str(url_index))
                    url = self.products_url[url_index]['url']

                    if (url == None) or (url == ""):  # If url doesn't exist, iteration is finished
                        self.work_finished = True
                        break

                    print(url)

                    product_info = {
                        'name': '',
                        'price': '',
                        'amount': '',
                        'unit': '',
                        'code': '',
                        'description': '',
                        'category': ''
                    }

                    req = Request(url, headers={
                        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36',
                        'cookie': 'ifoodV=02702425424564; ifoodUV=05110042592983; cto_lwid=0faa422c-9073-41d3-878c-67a029e4f472; ifoodV=04453954546708; ifoodUV=01912465359997; _gcl_au=1.1.1150787047.1568298592; _ga=GA1.3.1978835083.1568298594; _gid=GA1.3.299917675.1568298594; __kdtv=t%3D1568298595340%3Bi%3D2bdc535431d12da39a48db005a24b9e37b619776; _kdt=%7B%22t%22%3A1568298595340%2C%22i%22%3A%222bdc535431d12da39a48db005a24b9e37b619776%22%7D; ab.storage.deviceId.bf16c79a-2955-4975-a5c1-135e96fe4be4=%7B%22g%22%3A%224a9da739-3584-b3a7-63d3-7639c2b47329%22%2C%22c%22%3A1568298596075%2C%22l%22%3A1568298596075%7D; SL_C_23361dd035530_KEY=d4cc31315fbb7d56b564582c1667381a9380dc2c; _hjid=a14a97fb-c5ac-454b-95e1-414c039efb84; _fbp=fb.2.1568298598877.1476305068; sback_client=59710af9cba5a162024024e7; sback_partner=false; sb_days=1568298603328; intercom-id-qe4fnx70=a9fc1b6a-9b4e-4a85-8f2a-113e7b8ef719; sback_refresh_wp=no; _gaexp=GAX1.3.LBgZjeahSu2ceFVMR_7s3A.18241.0; ab.storage.userId.bf16c79a-2955-4975-a5c1-135e96fe4be4=%7B%22g%22%3A%22791046%22%2C%22c%22%3A1568299136255%2C%22l%22%3A1568299136255%7D; __kdtc=cid%3D791046%3Bt%3D1568299136216; ifoodShopCep=05448000; ifoodShopCidade=Rio+de+Janeiro; sback_pageview=false; _hjIncludedInSample=1; StandoutTag=51b0cb22-c0d9-6d7a-5a06-abb7b21aec06; _cm_ads_activation_retry=false; sback_browser=0-08792800-1568298611afa3367dc8a5676b228073cfb4557cad92061d4f12805493565d7a56731579f1-52206642-23226133156,13017628147-1568390601; sback_customer=$2wVxIUUlJjdOpGa6ZFcqhnTmNTaVhFVzllWqdjT102TOFWN30kStljT5gWSGJHR29kS3VXU2Q0daFGNy0UM6pWW2$12; sback_access_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJhcGkuc2JhY2sudGVjaCIsImlhdCI6MTU2ODM5MDYwMywiZXhwIjoxNTY4NDc3MDAzLCJhcGkiOiJ2MiIsImRhdGEiOnsiY2xpZW50X2lkIjoiNTk3MTBhZjljYmE1YTE2MjAyNDAyNGU3IiwiY2xpZW50X2RvbWFpbiI6InNob3AuaWZvb2QuY29tLmJyIiwiY3VzdG9tZXJfaWQiOiI1ZDdhNTY3NWFjYzY5YzZhMTgwNGQ4M2MiLCJjdXN0b21lcl9hbm9ueW1vdXMiOmZhbHNlLCJjb25uZWN0aW9uX2lkIjoiNWQ3YTU2NzVhY2M2OWM2YTE4MDRkODNkIiwiYWNjZXNzX2xldmVsIjoiY3VzdG9tZXIifX0.q_lFzuL6_A5F5p2WKv8KFT0V0xyJcLjVji6IDixBlB4.WrWrDruyiYKqHeqBuyqBKq; sback_customer_w=true; sback_session=5d7c5ac3c3d622b6ed6f275d; SL_C_23361dd035530_VID=YyCsCgztnU; SL_C_23361dd035530_SID=ImlUft4LTVX; JSESSIONID=444EAA4692F2EC7CA8B82EDE119F044C; _st_ses=6185966452654106; _st_no_script=1; _sptid=2471; _spcid=2409; _st_id=bmV3LmV5SjBlWEFpT2lKS1YxUWlMQ0poYkdjaU9pSklVekkxTmlKOS5leUpsYldGcGJDSTZJbWxyWVc1dmN5NWpiMjUwWVdOMFFHZHRZV2xzTG1OdmJTSjkudm5LNTVWYVZtS2JCTE9WWUMtbExncDdlUVVKdWVfc2dlTFJiZ0twa1MxUS5XcldyRHJ1eWlZelJLcURyelJ1eUhl; _st_idb=bmV3LmV5SjBlWEFpT2lKS1YxUWlMQ0poYkdjaU9pSklVekkxTmlKOS5leUpsYldGcGJDSTZJbWxyWVc1dmN5NWpiMjUwWVdOMFFHZHRZV2xzTG1OdmJTSjkudm5LNTVWYVZtS2JCTE9WWUMtbExncDdlUVVKdWVfc2dlTFJiZ0twa1MxUS5XcldyRHJ1eWlZelJLcURyelJ1eUhl; sback_current_session=1; sback_total_sessions=10; _spl_pv=283; ab.storage.sessionId.bf16c79a-2955-4975-a5c1-135e96fe4be4=%7B%22g%22%3A%22b2000598-9571-c54d-e171-358c41e5b46b%22%2C%22e%22%3A1568437273820%2C%22c%22%3A1568435460706%2C%22l%22%3A1568435473820%7D; intercom-session-qe4fnx70=U3JqakdwYnpYTEFUN0pld3VtVytWeGovc293QzllNTA4QnI0a1ZhenlyMU5jZ3BLWjlxUGFpM0ZFNzExMEJDKy0tR3NiTXh6RlUweGh1M3NtMmozcUdwQT09--afc4e294990c57ea4ea04413a0cd625688baaae3'})


                    page = urllib.request.urlopen(req)
                    soup = BeautifulSoup(page, 'html.parser')

                    # product name
                    try:
                        name = soup.find(class_="productName").text
                        product_info['name'] = ' '.join(name.split())
                    except:
                        product_info['name'] = ''

                    #product price
                    try:
                        price = soup.find(class_="activePrice").text
                        price = ' '.join(price.split())
                        product_info['price'] = price.replace('R$ ', '')
                    except:
                        product_info['price'] = ''

                    #product amount
                    try:
                        amount = soup.find(class_="amount text-mediumgrey").text
                        amount = re.findall(r'[-+]?\d*\.\d+|\d+', amount)
                        product_info['amount'] = amount[0]
                    except:
                        product_info['amount'] = ''

                    #product price per unit
                    try:
                        priceperunit = soup.find(class_="pricePerUnit").text
                        priceperunit = ' '.join(priceperunit.split())
                        product_info['unit'] = priceperunit.replace('R$ ', '')
                    except:
                        product_info['unit'] = ''

                    #description
                    try:
                        desc_soup = soup.find_all(class_="descriptionWrapper")
                        description = desc_soup[0].text
                        product_info['description'] = ' '.join(description.split())
                    except:
                        product_info['description'] = ''

                    # code
                    try:
                        code = desc_soup[1].text
                        code = ' '.join(code.split()).replace('Código: ', '')
                        product_info['code'] = code
                    except:
                        product_info['code']

                    # category
                    product_info['category'] = self.products_url[url_index]['category']

                    #output google sheet
                    while True:
                        try:
                            cell_list = self.sheet.range(url_index+2, 2, url_index+2, 8)
                            cell_values = [product_info['name'], product_info['category'], product_info['price'], product_info['unit'], product_info['amount'], product_info['code'], product_info['description']]
                            for i, val in enumerate(cell_values):  # gives us a tuple of an index and value
                                cell_list[i].value = val  # use the index on cell_list and the val from cell_values

                            self.sheet.update_cells(cell_list)
                            url_index += 1
                        except:
                            if self.credentials.access_token_expired:
                                print("Here is except token expired")
                                self.client.login()  # refreshes the token
                                continue
                            else:
                                print("normal except")
                                continue
                        break

                    print(product_info)
                except:
                    print("There is a issue, but this will be checked ASAP!")
                    continue
                break

            if (self.work_finished):
                break

        self.browser.close()

if __name__ == "__main__":

    ExtractProduct = ExtractProduct()
    ExtractProduct.extract_each_product_url_and_save()
    print("Task completed")