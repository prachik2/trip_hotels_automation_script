# -*- coding: utf-8 -*-
import scrapy
import time
import xlrd
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class TripHotels(scrapy.Spider):
    name = 'trip_hotels'
    allowed_domains = ['trip.com']

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        global start_url_list
        start_url_list = []
        loc = "/TripReport.xlsx"

        driver = webdriver.Chrome()

        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(1)

        sheet.cell_value(0, 0)
        row_data_list = []
        row_data_dict = {}
        for i in range(sheet.nrows - 1):
            row_data_list.append(sheet.row_values(i + 1))
        for j in row_data_list:
            row_data_dict["website"] = j[0]
            row_data_dict["destination"] = j[1]
            row_data_dict["check_in_date"] = xlrd.xldate.xldate_as_datetime(j[2], wb.datemode).date()
            row_data_dict["check_out_date"] = xlrd.xldate.xldate_as_datetime(j[3], wb.datemode).date()
            row_data_dict["room"] = int(j[4])
            row_data_dict["adult"] = int(j[5])
            row_data_dict["children"] = int(j[6])

            driver.get("https://www." + str(row_data_dict["website"]).lstrip())

            driver.maximize_window()

            wait = WebDriverWait(driver, 5)
            driver.implicitly_wait(8)
            # WebDriverWait(driver, 5)
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="hotels-destination"]')))
            destination_box = driver.find_element_by_xpath('//*[@id="hotels-destination"]')
            # WebDriverWait(driver, 20)
            # destination_box.clear()
            destination_box.send_keys(row_data_dict["destination"])

            checkin = driver.find_element_by_xpath(
                '/html/body/div[1]/div[4]/div[3]/div/div[2]/div/div[2]/div/div/div/ul/li[2]/div/div[1]/input').click()
            select_month_year = driver.find_element_by_xpath(
                '/html/body/div[1]/div[4]/div[3]/div/div[2]/div/div[2]/div/div/div/ul/li[2]/div/div[4]/div/div['
                '1]/div[1]/h3')
            # select_month_year.send_keys("Nov 2020")
            WebDriverWait(driver, 10)
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'search-btn-wrap')))
            try:
                search_icon = driver.find_element(By.CLASS_NAME, 'search-btn-wrap')
                search_icon.click()
            except exceptions.NoSuchElementException as e:
                pass
            wait = WebDriverWait(driver, 10)
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'long-list')))
            try:
                ul_list = driver.find_element_by_class_name('long-list')
                wait.until(EC.presence_of_element_located((By.TAG_NAME, 'li')))
                all_li = ul_list.find_elements_by_tag_name("li")
                for item in all_li:
                    try:
                        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'span')))
                        if item.find_element_by_tag_name("span") and item.is_displayed():
                            WebDriverWait(driver, 20)
                            # get the window handle after the window has opened
                            window_before = driver.window_handles[0]

                            window_before_title = driver.title

                            new_page = item.find_element_by_tag_name("span")
                            new_page.click()

                            # get the window handle after a new window has opened
                            window_after = driver.window_handles[0]

                            # switch on to new child window
                            driver.switch_to.window(window_after)
                            time.sleep(10)

                            window_after_title = driver.title
                            winnn = driver.current_url

                            # Compare and verify that main window and child window title don't match
                            if window_before_title != window_after_title:
                                print('Context switched to Twitter, so the title did not match')
                            else:
                                print('Control did not switch to new window')
                    except exceptions.NoSuchElementException as e:
                        print(e)
                        pass

                    except exceptions.StaleElementReferenceException as e:
                        print(e)
                        pass

            except exceptions.NoSuchElementException as e:
                pass

    def start_requests(self):
        for url in start_url_list:
            yield scrapy.Request(url, self.parse)

    def parse(self, response):
        for item in response.css('card-item-wrap'):
            data = {
                'source_url': item.url,
                'hotel_name': item.css('name').extract_first(),
            }
            count = 0
            for rate in item.css('more-star-repeat'):
                count += 1

            data['rate'] = count

            # callback = functools.partial(self.parse_article_page, data)
            # if data["manager_url"]:
            #     yield scrapy.Request(url=data["manager_url"], callback=callback)
            # else:
            #     pass
