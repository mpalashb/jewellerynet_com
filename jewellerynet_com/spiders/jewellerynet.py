# -*- coding: utf-8 -*-
from time import sleep
from scrapy import Spider
from selenium import webdriver
from scrapy.selector import Selector
from scrapy.http import Request
from selenium.common.exceptions import NoSuchElementException
import requests
from lxml.html import fromstring
import os
import csv
import glob
from openpyxl import Workbook


class JewellerynetSpider(Spider):
    name = 'jewellerynet'
    url = "https://exhibitors.jewellerynet.com/September-Hong-Kong-Jewellery-and-Gem-Fair-2019/"

    base_dict = []
    no_pos = 0

    def start_requests(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option(
            "excludeSwitches", ["ignore-certificate-errors"])
        self.options.add_argument('--disable-gpu')
        self.options.add_argument('--headless')

        path = r""

        self.driver = webdriver.Chrome(
            chrome_options=self.options, executable_path=path)
        # self.driver = webdriver.Chrome(path)
        self.driver.get(self.url)
        print("5 sec")
        sleep(5)
        # show_100    = self.driver.find_element_by_xpath('//*[@name="dtSearch_length"]//option[4]')
        # if show_100:
        #     show_100.click()
        #     print("7 sec--")
        #     sleep(7)

        sel = Selector(text=self.driver.page_source)
        elements = sel.xpath('//table//tr[@role="row"]')
        if elements:
            elements = elements[1:]
            for selector in elements:
                link = selector.xpath(
                    './/*[@class="ExhList_CompanyName"]//@href').extract_first()

                Name = selector.xpath(
                    './/*[@class="ExhList_CompanyName"]//text()').extract_first()

                Stand_No = selector.xpath(
                    './/*[@class="ExhList_BoothNo"]//text()').extract()
                if Stand_No:
                    Stand_No = ' '.join(Stand_No)

                Country_Region = selector.xpath(
                    './/td[4]//text()').extract_first()

                Product_category = selector.xpath('.//td[5]//text()').extract()
                if Product_category:
                    Product_category = ' '.join(Product_category)

                str_dict = {
                    'link': link,
                    'Name': Name,
                    'Stand_No': Stand_No,
                    'Country_Region': Country_Region,
                    'Product_category': Product_category}

        page_x = 0

        while True:
            try:
                page_x += 1
                # next_page=self.driver.find_element_by_xpath('//*[@class="paginate_button next"]')
                # next_page=self.driver.find_element_by_xpath('//*[contains(text(),"Next")]')
                # next_page=self.driver.find_element_by_xpath('//*[@class="paginate_button "]//*[contains(text(),"2")]')

                self.driver.find_element_by_xpath(
                    '//*[@class="paginate_button next"]').click()
                sleep(15)
                print('Sleeping 15 sec')

                print('\n')
                print('----------------------')
                print('----------------------')
                print('----------------------')

                sel = Selector(text=self.driver.page_source)
                elements = sel.xpath('//table//tr[@role="row"]')
                if elements:
                    elements = elements[1:]
                    for selector in elements:
                        link = selector.xpath(
                            './/*[@class="ExhList_CompanyName"]//@href').extract_first()

                        Name = selector.xpath(
                            './/*[@class="ExhList_CompanyName"]//text()').extract_first()

                        Stand_No = selector.xpath(
                            './/*[@class="ExhList_BoothNo"]//text()').extract()
                        if Stand_No:
                            Stand_No = ' '.join(Stand_No)

                        Country_Region = selector.xpath(
                            './/td[4]//text()').extract_first()

                        Product_category = selector.xpath(
                            './/td[5]//text()').extract()
                        if Product_category:
                            Product_category = ' '.join(Product_category)

                        str_dict = {
                            'link': link,
                            'Name': Name,
                            'Stand_No': Stand_No,
                            'Country_Region': Country_Region,
                            'Product_category': Product_category}

                        if link:
                            yield Request(link, meta=str_dict, callback=self.parse_p)
                        else:
                            yield Request("http://quotes.toscrape.com/", meta=str_dict, dont_filter=True, callback=self.parse_p)

            except NoSuchElementException:
                self.logger.info('No more pages to load.')
                self.driver.quit()
                break

    def parse_p(self, response):
        link = response.meta['link']
        Name = response.meta['Name']
        Stand_No = response.meta['Stand_No']
        Country_Region = response.meta['Country_Region']
        Product_category = response.meta['Product_category']
        if response.url == "http://quotes.toscrape.com/":
            Website = "None"
            Site_Response = "None"
            Site_Title = "None"

        Website = response.xpath(
            '//dt[contains(text(),"Our Website")]//following-sibling::dd//@href').extract_first()
        if Website is None:
            Website = "None"
        Site_Response = "None"
        Site_Title = ""

        try:
            r = requests.get(Website)
            Site_Response = r.status_code
            tree = fromstring(r.content)
            Site_Title = tree.findtext('.//title')
            if Site_Title:
                Site_Title = Site_Title.strip()
        except Exception as p:
            print(p)

        Emeralds = None
        Jewellery = None
        Diamond = None
        Rubies = None
        White_diamond = None
        Fancy_coloured_diamonds = None

        if "Emeralds" in str(Product_category):
            Emeralds = "Yes"

        if "Jewellery" in str(Product_category):
            Jewellery = "Yes"

        if "Diamond" in str(Product_category):
            Diamond = "Yes"

        if "Rubies" in str(Product_category):
            Rubies = "Yes"

        if "White diamond" in str(Product_category):
            White_diamond = "Yes"

        if "Fancy coloured diamonds" in str(Product_category):
            Fancy_coloured_diamonds = "Yes"

        if Website == "http://":
            Website = "None"

        yield{

            'link': link,
            'Name': Name,
            'Stand_No': Stand_No,
            'Country_Region': Country_Region,
            'Product_category': Product_category,
            'Website': Website,
            'Site_Response': Site_Response,
            'Site_Title': Site_Title,
            "Emeralds": Emeralds,
            "Jewellery": Jewellery,
            "Diamond": Diamond,
            "Rubies": Rubies,
            "White_diamond": White_diamond,
            "Fancy_coloured_diamonds": Fancy_coloured_diamonds

        }

        self.no_pos += 1
        print('---------------------')
        print(f"{self.no_pos} done!")
        print('---------------------')

    def close(self, reason):
        csv_file = max(glob.iglob('*.csv'), key=os.path.getctime)

        wb = Workbook()
        ws = wb.active

        with open(csv_file, 'r', encoding='utf-8') as f:
            for row in csv.reader(f):
                ws.append(row)

        wb.save(csv_file.replace('.csv', '') + '.xlsx')
